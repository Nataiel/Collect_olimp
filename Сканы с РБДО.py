import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import requests
from pathlib import Path
from urllib.parse import urlparse
import threading
import queue
import re


class PDFDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📥 Скачивание PDF файлов")
        self.root.geometry("650x550")
        self.root.resizable(True, True)
        self.root.minsize(500, 400)

        self.file_path = tk.StringVar()
        self.dir_path = tk.StringVar()

        self.download_queue = queue.Queue()
        self.download_thread = None

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(file_frame, text="Файл со ссылками (.txt):").pack(side=tk.LEFT)
        ttk.Entry(file_frame, textvariable=self.file_path, state="readonly", width=50).pack(side=tk.LEFT, padx=5,
                                                                                            fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Обзор...", command=self.select_file).pack(side=tk.LEFT)

        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        ttk.Label(dir_frame, text="Директория сохранения:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.dir_path, state="readonly", width=50).pack(side=tk.LEFT, padx=5,
                                                                                          fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="Обзор...", command=self.select_dir).pack(side=tk.LEFT)

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=15)
        self.start_btn = ttk.Button(control_frame, text="🚀 Начать загрузку", command=self.start_download)
        self.start_btn.pack(fill=tk.X)
        self.progress = ttk.Progressbar(control_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(10, 0))

        log_frame = ttk.LabelFrame(main_frame, text="Журнал операций")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите текстовый файл со ссылками",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)

    def select_dir(self):
        dirname = filedialog.askdirectory(title="Выберите папку для сохранения PDF")
        if dirname:
            self.dir_path.set(dirname)

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    @staticmethod
    def extract_13digit_code(url: str) -> str | None:
        """
        Извлекает из URL первую последовательность ровно из 13 цифр.
        Например, для ссылки .../9007538770086/... вернёт '9007538770086'.
        """
        # (?<!\d) и (?!\d) гарантируют, что мы ищем ровно 13 цифр,
        # а не часть более длинного числа
        match = re.search(r'(?<!\d)\d{13}(?!\d)', url)
        return match.group(0) if match else None

    def start_download(self):
        file_path = self.file_path.get()
        dir_path = self.dir_path.get()

        if not file_path or not Path(file_path).exists():
            self.log("❌ Ошибка: Пожалуйста, выберите существующий файл со ссылками.")
            return
        if not dir_path:
            self.log("❌ Ошибка: Пожалуйста, выберите директорию для сохранения.")
            return

        self.start_btn.config(state=tk.DISABLED, text="⏳ Загрузка...")
        self.progress['value'] = 0
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log("📁 Подготовка к загрузке...")

        self.download_thread = threading.Thread(
            target=self.download_worker,
            args=(file_path, dir_path),
            daemon=True
        )
        self.download_thread.start()
        self.root.after(100, self.check_queue)

    def download_worker(self, file_path, dir_path):
        try:
            download_path = Path(dir_path)
            download_path.mkdir(parents=True, exist_ok=True)

            with open(file_path, 'r', encoding='utf-8') as f:
                urls = [line.strip() for line in f if line.strip() and not line.startswith('#')]

            total_urls = len(urls)
            self.download_queue.put(f"🔗 Найдено ссылок для загрузки: {total_urls}\n")

            for index, url in enumerate(urls, 1):
                self.download_queue.put(f"⏳ [{index}/{total_urls}] Загрузка: {url}")

                try:
                    response = requests.get(url, stream=True, timeout=15)
                    response.raise_for_status()

                    # === ЛОГИКА ИМЕНОВАНИЯ ФАЙЛА ===
                    code = self.extract_13digit_code(url)
                    if code:
                        filename = f"{code}.pdf"
                    else:
                        # Fallback: если 13-значного кода нет, берём имя из URL
                        parsed_url = urlparse(url)
                        filename = Path(parsed_url.path).name or "downloaded_document.pdf"
                        self.download_queue.put(f"⚠️ 13-значный код не найден, использовано имя из URL: {filename}")
                    # ================================

                    filepath = download_path / filename

                    with open(filepath, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)

                    self.download_queue.put(f"✅ Сохранено как: {filename}")

                except requests.exceptions.RequestException as e:
                    self.download_queue.put(f"❌ Ошибка при загрузке {url}: {e}")
                except Exception as e:
                    self.download_queue.put(f"❌ Непредвиденная ошибка для {url}: {e}")

                progress_percent = int((index / total_urls) * 100)
                self.download_queue.put(("PROGRESS", progress_percent))

        except Exception as e:
            self.download_queue.put(f"❌ Критическая ошибка: {e}")
        finally:
            self.download_queue.put("FINISHED")

    def check_queue(self):
        try:
            while True:
                msg = self.download_queue.get_nowait()

                if msg == "FINISHED":
                    self.start_btn.config(state=tk.NORMAL, text="🚀 Начать загрузку")
                    self.log("🎉 Загрузка завершена!")
                    return

                if isinstance(msg, tuple) and msg[0] == "PROGRESS":
                    self.progress['value'] = msg[1]
                else:
                    self.log(str(msg))

        except queue.Empty:
            pass

        if self.download_thread and self.download_thread.is_alive():
            self.root.after(100, self.check_queue)


if __name__ == "__main__":
    try:
        root = tk.Tk()
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        app = PDFDownloaderApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка запуска GUI: {e}")