import os
import sys
import subprocess
import zipfile
import rarfile
import tarfile
import gzip
import bz2
import lzma
import time
import multiprocessing
import shutil
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, QTextEdit,
    QHBoxLayout, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# Centralize extensões suportadas
ARCHIVE_EXTENSIONS = (
    '.zip', '.rar', '.7z',
    '.tar', '.gz', '.bz2', '.xz',
    '.tgz', '.tbz2', '.txz',
    '.tar.gz', '.tar.bz2', '.tar.xz'
)

APP_STYLE = """
QWidget { background-color: #2D2D2D; color: #E0E0E0; font-family: 'Segoe UI'; font-size: 12px; border: none; }
QMainWindow { background-color: #2D2D2D; border: 1px solid #444; }
QPushButton { background-color: #3A3A3A; border: 1px solid #444; border-radius: 4px; padding: 8px 16px; min-width: 100px; color: #E0E0E0; }
QPushButton:hover { background-color: #4A4A4A; }
QPushButton:pressed { background-color: #2A2A2A; }
QPushButton:disabled { background-color: #2A2A2A; color: #777; }
QLineEdit { background-color: #3A3A3A; border: 1px solid #444; border-radius: 4px; padding: 8px; selection-background-color: #505050; }
QTextEdit { background-color: #3A3A3A; border: 1px solid #444; border-radius: 4px; padding: 8px; font-family: 'Consolas', monospace; }
QLabel { color: #E0E0E0; padding: 4px 0; }
QStatusBar { background-color: #3A3A3A; border-top: 1px solid #444; padding: 4px; }
QProgressBar { border: 1px solid #444; border-radius: 4px; text-align: center; background-color: #3A3A3A; }
QProgressBar::chunk { background-color: #4A6FA5; width: 10px; }
"""

class ExtractionThread(QThread):
    update_progress = pyqtSignal(int, str)
    update_status = pyqtSignal(str)
    extraction_done = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.root_folder = ""
        self.password = ""
        self.executor = ThreadPoolExecutor(max_workers=multiprocessing.cpu_count())
        self._is_running = True

    def find_archive_folders(self):
        """Retorna pastas que possuem arquivos compactados."""
        return [
            root for root, _, files in os.walk(self.root_folder)
            if any(f.lower().endswith(ARCHIVE_EXTENSIONS) for f in files)
        ]

    def find_latest_archive(self, folder):
        """Retorna o arquivo compactado mais recente em uma pasta."""
        valid_files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(ARCHIVE_EXTENSIONS)
        ]
        if not valid_files:
            return None
        return max(valid_files, key=os.path.getmtime)

    def extract_7z(self, archive_path, output_folder, password):
        """Extração otimizada usando 7-Zip via subprocess."""
        seven_zip = "C:\\Program Files\\7-Zip\\7z.exe"
        if not os.path.exists(seven_zip):
            raise Exception("7-Zip não encontrado em C:\\Program Files\\7-Zip\\7z.exe")
        cmd = [
            seven_zip, "x", archive_path,
            f"-o{output_folder}", "-y", "-mmt=on", "-bb3", "-sccUTF-8"
        ]
        if password:
            cmd.extend([f"-p{password}", "-mhe=on"])
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            stdin=subprocess.PIPE,
            creationflags=(
                subprocess.CREATE_NO_WINDOW |
                subprocess.HIGH_PRIORITY_CLASS
            ),
            text=True
        )

        # Monitoramento de progresso em thread separada
        import threading

        def monitor_progress():
            last_size = 0
            while process.poll() is None:
                try:
                    current_size = sum(
                        os.path.getsize(os.path.join(root, f))
                        for root, _, files in os.walk(output_folder)
                        for f in files
                    )
                    speed = (current_size - last_size) / (1024 * 1024)
                    last_size = current_size
                    self.update_status.emit(
                        f"Extraindo... {current_size/(1024*1024):.1f}MB ({speed:.1f} MB/s)"
                    )
                except Exception:
                    pass
                time.sleep(0.5)

        t = threading.Thread(target=monitor_progress, daemon=True)
        t.start()

        # Log de progresso
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                print(output.strip())

        t.join(timeout=0.1)

        if process.returncode != 0:
            error_msg = process.stderr.read()
            raise Exception(f"Erro {process.returncode}: {error_msg}")

        return True

    def extract_zip(self, archive_path, output_folder, password):
        """Extração otimizada para ZIP."""
        with zipfile.ZipFile(archive_path) as zf:
            for file in zf.infolist():
                target_path = os.path.join(output_folder, file.filename)
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                with open(target_path, 'wb') as f:
                    shutil.copyfileobj(
                        zf.open(file, pwd=password.encode() if password else None), f
                    )

    def extract_rar(self, archive_path, output_folder, password):
        """Extração paralela para RAR."""
        with rarfile.RarFile(archive_path) as rf:
            file_list = rf.infolist()
            def extract_file(file):
                if not self._is_running:
                    return
                rf.extract(file, path=output_folder, pwd=password)
            list(self.executor.map(extract_file, file_list))

    def extract_tar(self, archive_path, output_folder):
        """Extração para TAR e derivados."""
        ext = archive_path.lower()
        if ext.endswith('.tar'):
            mode = 'r'
        elif ext.endswith('.tar.gz'):
            mode = 'r:gz'
        elif ext.endswith('.tar.bz2'):
            mode = 'r:bz2'
        elif ext.endswith('.tar.xz'):
            mode = 'r:xz'
        else:
            mode = 'r'
        with tarfile.open(archive_path, mode) as tf:
            tf.extractall(path=output_folder)

    def extract_simple(self, archive_path, output_folder, ext):
        """Extração para GZ, BZ2, XZ, TGZ, TBZ2, TXZ."""
        openers = {
            '.gz': gzip.open,
            '.tgz': gzip.open,
            '.bz2': bz2.open,
            '.tbz2': bz2.open,
            '.xz': lzma.open,
            '.txz': lzma.open
        }
        opener = openers.get(ext)
        if opener:
            with opener(archive_path, 'rb') as f_in:
                out_name = os.path.splitext(os.path.basename(archive_path))[0]
                with open(os.path.join(output_folder, out_name), 'wb') as f_out:
                    f_out.write(f_in.read())

    def extract_archive(self, archive_path, output_base):
        """Seleciona o método de extração apropriado."""
        try:
            archive_name = os.path.basename(archive_path)
            output_folder = os.path.join(output_base, "extracted_" + os.path.splitext(archive_name)[0])
            os.makedirs(output_folder, exist_ok=True)
            original_size = os.path.getsize(archive_path) / (1024 * 1024)
            ext = archive_path.lower()

            if ext.endswith('.7z'):
                self.update_status.emit("Extraindo com 7-Zip (máximo desempenho)...")
                self.extract_7z(archive_path, output_folder, self.password)
            elif ext.endswith('.zip'):
                self.extract_zip(archive_path, output_folder, self.password)
            elif ext.endswith('.rar'):
                self.extract_rar(archive_path, output_folder, self.password)
            elif ext.endswith(('.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                self.extract_tar(archive_path, output_folder)
            elif ext.endswith(('.gz', '.tgz', '.bz2', '.tbz2', '.xz', '.txz')):
                for e in ['.gz', '.tgz', '.bz2', '.tbz2', '.xz', '.txz']:
                    if ext.endswith(e):
                        self.extract_simple(archive_path, output_folder, e)
                        break
            else:
                raise Exception("Formato não suportado.")

            extracted_size = sum(
                os.path.getsize(os.path.join(root, f))
                for root, _, files in os.walk(output_folder)
                for f in files
            ) / (1024 * 1024)

            return {
                "status": "Sucesso",
                "message": f"Extraído via {'Otimizado' if ext.endswith(('.zip', '.rar', '.7z')) else 'Python'}",
                "original_size_mb": round(original_size, 2),
                "extracted_size_mb": round(extracted_size, 2),
                "files": os.listdir(output_folder)
            }

        except Exception as e:
            return {
                "status": "Erro",
                "message": str(e),
                "original_size_mb": round(original_size, 2) if 'original_size' in locals() else 0,
                "extracted_size_mb": 0,
                "files": []
            }

    def run(self):
        self._is_running = True
        archive_folders = self.find_archive_folders()
        total_results = {}
        start_time = time.time()

        if not archive_folders:
            self.update_status.emit("Nenhum arquivo encontrado")
            self.extraction_done.emit({})
            return

        total_files = len(archive_folders)

        for i, folder in enumerate(archive_folders):
            if not self._is_running:
                break
            folder_start_time = time.time()
            self.update_status.emit(f"Processando: {os.path.basename(folder)}...")

            latest_file = self.find_latest_archive(folder)
            if not latest_file:
                total_results[folder] = {"status": "Ignorado", "message": "Nenhum arquivo válido"}
                continue

            elapsed = time.time() - start_time
            avg_time_per_file = elapsed / (i + 1e-6)
            remaining_files = total_files - (i + 1)
            remaining_time = avg_time_per_file * remaining_files

            if remaining_time > 3600:
                time_str = f"{remaining_time/3600:.1f} horas restantes"
            elif remaining_time > 60:
                time_str = f"{remaining_time/60:.1f} minutos restantes"
            else:
                time_str = f"{remaining_time:.0f} segundos restantes"

            progress = int((i + 1) / total_files * 100)
            self.update_progress.emit(progress, time_str)

            result = self.extract_archive(latest_file, folder)
            folder_time = time.time() - folder_start_time
            result["processing_time"] = f"{folder_time:.1f}s"
            total_results[folder] = result

        self.extraction_done.emit(total_results)

    def stop(self):
        self._is_running = False
        self.executor.shutdown(wait=False)

class BackupExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Extrator de Arquivos Automático")
        self.setGeometry(200, 200, 900, 700)
        self.initUI()

    def initUI(self):
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title = QLabel("📦 Extrator de Arquivos Automático")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title, alignment=Qt.AlignCenter)

        select_group = QWidget()
        select_layout = QHBoxLayout()
        self.folder_label = QLabel("📂 Pasta principal: Não selecionada")
        select_layout.addWidget(self.folder_label, stretch=4)
        self.select_btn = QPushButton("Selecionar Pasta")
        self.select_btn.setStyleSheet("background-color: #4A6FA5;")
        self.select_btn.clicked.connect(self.select_root_folder)
        select_layout.addWidget(self.select_btn, stretch=1)
        select_group.setLayout(select_layout)
        layout.addWidget(select_group)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("🔑 Digite a senha (deixe em branco se não tiver)")
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.password_input)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.time_label = QLabel("Tempo restante: calculando...")
        layout.addWidget(self.time_label)

        self.extract_btn = QPushButton("▶ Extrair Arquivos Recentes")
        self.extract_btn.setStyleSheet("background-color: #5A8E5A; font-weight: bold;")
        self.extract_btn.clicked.connect(self.start_extraction)
        self.extract_btn.setEnabled(False)
        layout.addWidget(self.extract_btn)

        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        layout.addWidget(self.report_text, stretch=1)

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Pronto")

        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

    def select_root_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta Principal")
        if folder:
            self.root_folder = folder
            self.folder_label.setText(f"📂 Pasta principal: {folder}")
            self.extract_btn.setEnabled(True)
            self.status_bar.showMessage(f"Pasta selecionada: {os.path.basename(folder)}")

    def start_extraction(self):
        self.thread = ExtractionThread()
        self.thread.root_folder = self.root_folder
        self.thread.password = self.password_input.text()

        self.thread.update_progress.connect(self.update_progress)
        self.thread.update_status.connect(self.update_status)
        self.thread.extraction_done.connect(self.extraction_complete)

        self.set_ui_enabled(False)
        self.report_text.clear()
        self.progress_bar.setValue(0)
        self.thread.start()

    def update_progress(self, value, time_remaining):
        self.progress_bar.setValue(value)
        self.time_label.setText(f"⏳ Tempo estimado: {time_remaining}")
        if value % 5 == 0 or value < 10:
            QApplication.processEvents()

    def update_status(self, message):
        self.status_bar.showMessage(message)

    def extraction_complete(self, results):
        self.set_ui_enabled(True)
        self.generate_report(results)
        errors = sum(1 for r in results.values() if r['status'] == 'Erro')
        if errors:
            QMessageBox.warning(self, "Conclusão", f"Processo completo com {errors} erro(s)")
        else:
            QMessageBox.information(self, "Conclusão", "Extração concluída com sucesso!")

    def set_ui_enabled(self, enabled):
        self.select_btn.setEnabled(enabled)
        self.password_input.setEnabled(enabled)
        self.extract_btn.setEnabled(enabled)

    def generate_report(self, results):
        report_lines = [
            f"📝 RELATÓRIO DE EXTRAÇÃO - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"📂 Pasta principal: {self.root_folder}",
            f"🔑 Senha usada: {'Sim' if self.password_input.text() else 'Não'}",
            "\n" + "="*80
        ]

        total_original_size = 0
        total_extracted_size = 0

        for folder, data in results.items():
            folder_name = os.path.basename(folder)
            status_icon = "✅" if data['status'] == 'Sucesso' else "❌"

            report_lines.extend([
                f"\n📁 {folder_name}",
                f"   {status_icon} Status: {data['status']}",
                f"   📦 Tamanho original: {data.get('original_size_mb', 0):.2f} MB",
                f"   🗃️ Tamanho extraído: {data.get('extracted_size_mb', 0):.2f} MB",
                f"   💬 Mensagem: {data['message']}",
                f"   ⏱️ Tempo de processamento: {data.get('processing_time', 'N/A')}"
            ])

            if data.get('files'):
                report_lines.append("   📄 Arquivos extraídos:")
                report_lines.extend(f"      - {file}" for file in data['files'])

            total_original_size += data.get('original_size_mb', 0)
            total_extracted_size += data.get('extracted_size_mb', 0)

        report_lines.extend([
            "\n" + "="*80,
            f"ℹ️ Total de pastas processadas: {len(results)}",
            f"📊 Tamanho total original: {total_original_size:.2f} MB",
            f"📦 Tamanho total extraído: {total_extracted_size:.2f} MB"
        ])
        self.report_text.setPlainText("\n".join(report_lines))

        report_file = os.path.join(self.root_folder, "relatorio_extracao.txt")
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("\n".join(report_lines))
        self.status_bar.showMessage(f"Relatório salvo em: {report_file}", 5000)

if __name__ == "__main__":
    try:
        import rarfile
    except ImportError:
        print("Erro: Instale rarfile com: pip install rarfile")
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    window = BackupExtractor()
    window.show()
    sys.exit(app.exec_())