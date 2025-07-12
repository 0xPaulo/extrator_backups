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
    QHBoxLayout, QProgressBar, QListWidget, QListWidgetItem,
    QTabWidget, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

# Centralize extensões suportadas
ARCHIVE_EXTENSIONS = (
    '.zip', '.rar', '.7z',
    '.tar', '.gz', '.bz2', '.xz',
    '.tgz', '.tbz2', '.txz',
    '.tar.gz', '.tar.bz2', '.tar.xz'
)

APP_STYLE = """
/* Main Window Styling */
QMainWindow {
    background-color: #2D2D2D;
    border: none;
}

/* General Widget Styling */
QWidget {
    background-color: #2D2D2D;
    color: #E0E0E0;
    font-family: 'Segoe UI';
    font-size: 13px;
    border: none;
}

/* Tab Widget Styling */
QTabWidget::pane {
    border: 1px solid #444;
    border-radius: 4px;
    margin-top: 10px;
    background: #353535;
}

QTabBar::tab {
    background: #3A3A3A;
    color: #E0E0E0;
    padding: 10px 20px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    border: 1px solid #444;
    margin-right: 2px;
    min-width: 120px;
    font-size: 13px;
}

QTabBar::tab:selected {
    background: #4A6FA5;
    border-bottom-color: #4A6FA5;
}

QTabBar::tab:hover {
    background: #505050;
}

/* Button Styling */
QPushButton {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px 16px;
    min-width: 100px;
    color: #E0E0E0;
}

QPushButton:hover {
    background-color: #4A4A4A;
}

QPushButton:pressed {
    background-color: #2A2A2A;
}

QPushButton:disabled {
    background-color: #2A2A2A;
    color: #777;
}

/* Special Buttons */
QPushButton#extractButton {
    background-color: #5A8E5A;
    font-weight: bold;
    font-size: 15px;
    padding: 12px 30px;
    border: 1px solid #4A7A4A;
}

QPushButton#extractButton:hover {
    background-color: #6B9E6B;
}

QPushButton#extractButton:pressed {
    background-color: #4A7A4A;
}

QPushButton#folderButton {
    background-color: #4A6FA5;
    font-size: 13px;
    padding: 8px 12px;
}

/* Line Edit Styling */
QLineEdit {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
    selection-background-color: #505050;
    min-height: 32px;
}

QLineEdit:focus {
    border: 1px solid #4A6FA5;
}

/* Text Edit Styling */
QTextEdit {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
    font-family: 'Consolas', monospace;
    font-size: 12px;
    color: #E0E0E0;
}

/* Label Styling */
QLabel {
    color: #E0E0E0;
    padding: 4px 0;
}

QLabel#titleLabel {
    font-size: 24px;
    font-weight: bold;
    margin-bottom: 16px;
    color: #E0E0E0;
}

/* Status Bar Styling */
QStatusBar {
    background-color: #3A3A3A;
    border-top: 1px solid #444;
    padding: 4px;
    font-size: 12px;
}

/* Progress Bar Styling */
QProgressBar {
    border: 1px solid #444;
    border-radius: 4px;
    text-align: center;
    background-color: #3A3A3A;
    height: 24px;
}

QProgressBar::chunk {
    background-color: #4A6FA5;
    border-radius: 3px;
}

/* List Widget Styling */
QListWidget {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    font-size: 13px;
    padding: 4px;
    outline: 0;
}

QListWidget::item {
    padding: 6px;
    border-bottom: 1px solid #444;
}

QListWidget::item:selected {
    background-color: #4A6FA5;
    color: white;
}

QListWidget::item:hover {
    background-color: #505050;
}

/* Separator Styling */
QFrame#separator {
    background-color: #444;
    max-height: 1px;
    min-height: 1px;
}
"""

class ExtractionThread(QThread):
    update_progress = pyqtSignal(int, str)
    update_status = pyqtSignal(str)
    extraction_done = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.root_folder = ""
        self.password = ""
        self.selected_folders = None
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
        """Extração otimizada usando 7-Zip via subprocess com progresso detalhado."""
        seven_zip = "C:\\Program Files\\7-Zip\\7z.exe"
        if not os.path.exists(seven_zip):
            raise Exception("7-Zip não encontrado em C:\\Program Files\\7-Zip\\7z.exe")
        cmd = [
            seven_zip, "x", archive_path,
            f"-o{output_folder}", "-y", "-mmt=on", "-bb1", "-sccUTF-8"
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

        import threading

        def monitor_progress():
            last_size = 0
            last_time = time.time()
            while process.poll() is None:
                try:
                    current_size = sum(
                        os.path.getsize(os.path.join(root, f))
                        for root, _, files in os.walk(output_folder)
                        for f in files
                    )
                    now = time.time()
                    elapsed = now - last_time
                    speed = (current_size - last_size) / (1024 * 1024) / elapsed if elapsed > 0 else 0
                    last_size = current_size
                    last_time = now
                    self.update_status.emit(
                        f"Extraindo... {current_size/(1024*1024):.1f}MB ({speed:.1f} MB/s)"
                    )
                except Exception:
                    pass
                time.sleep(0.5)

        t = threading.Thread(target=monitor_progress, daemon=True)
        t.start()

        # Progresso detalhado por arquivo
        total_files = 0
        extracted_files = 0
        file_names = []
        file_name = ""
        for line in process.stdout:
            if line.startswith("- "):  # 7z -bb1 mostra arquivos assim
                file_name = line.strip()[2:]
                file_names.append(file_name)
                total_files += 1
            elif line.startswith("Extracting  "):
                extracted_files += 1
                percent = int((extracted_files / max(total_files, 1)) * 100)
                self.update_progress.emit(percent, f"Extraindo: {file_name}")
                self.update_status.emit(f"Extraindo: {file_name}")

        t.join(timeout=0.1)

        # Aguarde o processo terminar e capture stderr
        _, stderr = process.communicate()
        if process.returncode != 0:
            err_msg = stderr.strip()
            # Verifica se o erro é de senha
            if "Wrong password" in err_msg or "Can not open encrypted archive" in err_msg or "Data Error" in err_msg:
                raise Exception("Arquivo protegido por senha. Por favor, informe a senha correta.")
            raise Exception(f"Erro {process.returncode}: {err_msg}")

        return True

    def extract_zip(self, archive_path, output_folder, password):
        """Extração otimizada para ZIP."""
        try:
            with zipfile.ZipFile(archive_path) as zf:
                for file in zf.infolist():
                    target_path = os.path.join(output_folder, file.filename)
                    os.makedirs(os.path.dirname(target_path), exist_ok=True)
                    with open(target_path, 'wb') as f:
                        shutil.copyfileobj(
                            zf.open(file, pwd=password.encode() if password else None), f
                        )
        except RuntimeError as e:
            if "password required" in str(e).lower() or "Bad password" in str(e):
                raise Exception("Arquivo ZIP protegido por senha. Por favor, informe a senha correta.")
            raise

    def extract_rar(self, archive_path, output_folder, password):
        """Extração paralela para RAR."""
        try:
            with rarfile.RarFile(archive_path) as rf:
                file_list = rf.infolist()
                def extract_file(file):
                    if not self._is_running:
                        return
                    rf.extract(file, path=output_folder, pwd=password)
                list(self.executor.map(extract_file, file_list))
        except rarfile.BadRarFile as e:
            if "password" in str(e).lower():
                raise Exception("Arquivo RAR protegido por senha. Por favor, informe a senha correta.")
            raise

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

            # NOVO: pegar data de criação e modificação do arquivo
            archive_ctime = os.path.getctime(archive_path)
            archive_mtime = os.path.getmtime(archive_path)
            archive_ctime_str = datetime.fromtimestamp(archive_ctime).strftime('%Y-%m-%d %H:%M:%S')
            archive_mtime_str = datetime.fromtimestamp(archive_mtime).strftime('%Y-%m-%d %H:%M:%S')

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
                "files": os.listdir(output_folder),
                "latest_archive": archive_name,
                "latest_archive_ctime": archive_ctime_str,
                "latest_archive_mtime": archive_mtime_str
            }

        except Exception as e:
            return {
                "status": "Erro",
                "message": str(e),
                "original_size_mb": round(original_size, 2) if 'original_size' in locals() else 0,
                "extracted_size_mb": 0,
                "files": [],
                "latest_archive": archive_name if 'archive_name' in locals() else "",
                "latest_archive_ctime": archive_ctime_str if 'archive_ctime_str' in locals() else "",
                "latest_archive_mtime": archive_mtime_str if 'archive_mtime_str' in locals() else ""
            }

    def run(self):
        self._is_running = True
        if self.selected_folders is not None:
            archive_folders = self.selected_folders
        else:
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
        self.setWindowTitle("Extrator de Backups")
        self.setWindowIcon(QIcon.fromTheme('archive-extract'))
        self.setGeometry(200, 200, 1000, 750)
        self.initUI()

    def initUI(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Título
        title = QLabel("📦 Extrator de Backups")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # Separador
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setObjectName("separator")
        main_layout.addWidget(separator)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        main_layout.addWidget(self.tabs, stretch=1)

        # Aba de seleção de pastas
        select_tab = QWidget()
        select_layout = QVBoxLayout()
        select_layout.setContentsMargins(15, 15, 15, 15)
        select_layout.setSpacing(15)
        select_tab.setLayout(select_layout)

        # Grupo seleção pasta
        folder_group = QWidget()
        folder_group_layout = QHBoxLayout()
        folder_group_layout.setContentsMargins(0, 0, 0, 0)
        folder_group_layout.setSpacing(10)
        
        self.folder_label = QLabel("📂 Pasta principal: Não selecionada")
        self.folder_label.setStyleSheet("font-size: 14px;")
        folder_group_layout.addWidget(self.folder_label, stretch=4)
        
        self.select_btn = QPushButton("Selecionar Pasta")
        self.select_btn.setObjectName("folderButton")
        self.select_btn.clicked.connect(self.select_root_folder)
        folder_group_layout.addWidget(self.select_btn, stretch=1)
        
        folder_group.setLayout(folder_group_layout)
        select_layout.addWidget(folder_group)

        # Lista de pastas
        self.folder_list = QListWidget()
        self.folder_list.setSelectionMode(QListWidget.MultiSelection)
        self.folder_list.setMinimumHeight(200)
        select_layout.addWidget(self.folder_list)
        self.folder_list.hide()

        # NOVO: duplo clique para abrir pasta extraída
        self.folder_list.itemDoubleClicked.connect(self.open_extracted_folder)

        # NOVO: botões de seleção e exclusão
        btn_select_all = QPushButton("Marcar Todos")
        btn_unselect_all = QPushButton("Desmarcar Todos")
        btn_delete = QPushButton("Excluir Extraídos")
        btn_select_all.clicked.connect(self.select_all_folders)
        btn_unselect_all.clicked.connect(self.unselect_all_folders)
        btn_delete.clicked.connect(self.delete_extracted_folders)

        btns = QHBoxLayout()
        btns.addWidget(btn_select_all)
        btns.addWidget(btn_unselect_all)
        btns.addWidget(btn_delete)
        select_layout.addLayout(btns)

        # Grupo senha
        password_group = QWidget()
        password_layout = QVBoxLayout()
        password_layout.setContentsMargins(0, 0, 0, 0)
        password_layout.setSpacing(5)
        
        password_label = QLabel("🔑 Senha (opcional):")
        password_label.setStyleSheet("font-size: 13px;")
        password_layout.addWidget(password_label)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Digite a senha se os arquivos estiverem protegidos")
        self.password_input.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(self.password_input)
        
        password_group.setLayout(password_layout)
        select_layout.addWidget(password_group)

        # Botão de extrair centralizado
        btn_container = QWidget()
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(0, 10, 0, 0)
        
        self.extract_btn = QPushButton("▶ Extrair Arquivos Recentes")
        self.extract_btn.setObjectName("extractButton")
        self.extract_btn.clicked.connect(self.start_extraction)
        self.extract_btn.setEnabled(False)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.extract_btn)
        btn_layout.addStretch()
        
        btn_container.setLayout(btn_layout)
        select_layout.addWidget(btn_container)

        select_layout.addStretch()
        self.tabs.addTab(select_tab, "🔍 Seleção de Pastas")

        # Aba de relatório/log
        report_tab = QWidget()
        report_layout = QVBoxLayout()
        report_layout.setContentsMargins(15, 15, 15, 15)
        report_layout.setSpacing(15)
        report_tab.setLayout(report_layout)

        # Grupo progresso
        progress_group = QWidget()
        progress_layout = QVBoxLayout()
        progress_layout.setContentsMargins(0, 0, 0, 0)
        progress_layout.setSpacing(5)
        
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        
        self.time_label = QLabel("Tempo restante: aguardando início...")
        self.time_label.setStyleSheet("font-size: 13px; color: #AAAAAA;")
        progress_layout.addWidget(self.time_label)
        
        progress_group.setLayout(progress_layout)
        report_layout.addWidget(progress_group)

        # Relatório
        report_label = QLabel("📋 Relatório de Extração:")
        report_label.setStyleSheet("font-size: 14px;")
        report_layout.addWidget(report_label)
        
        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        report_layout.addWidget(self.report_text, stretch=1)

        self.tabs.addTab(report_tab, "📊 Relatório")

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Pronto")

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def select_root_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta Principal")
        if folder:
            self.root_folder = folder
            self.folder_label.setText(f"📂 Pasta principal: {folder}")
            self.extract_btn.setEnabled(True)
            self.status_bar.showMessage(f"Pasta selecionada: {os.path.basename(folder)}")
            # Listar subpastas com arquivos compactados
            self.folder_list.clear()
            self.folder_list.show()
            # Correção: use uma instância temporária para buscar as pastas
            temp_thread = ExtractionThread()
            temp_thread.root_folder = folder
            archive_folders = temp_thread.find_archive_folders()
            for subfolder in archive_folders:
                item = QListWidgetItem(subfolder)
                item.setCheckState(Qt.Checked)
                self.folder_list.addItem(item)

    def start_extraction(self):
        selected_folders = [
            self.folder_list.item(i).text()
            for i in range(self.folder_list.count())
            if self.folder_list.item(i).checkState() == Qt.Checked
        ]
        if not selected_folders:
            QMessageBox.warning(self, "Aviso", "Selecione ao menos uma pasta para extrair.")
            return

        self.thread = ExtractionThread()
        self.thread.root_folder = self.root_folder
        self.thread.password = self.password_input.text()
        self.thread.selected_folders = selected_folders

        self.thread.update_progress.connect(self.update_progress)
        self.thread.update_status.connect(self.update_status)
        self.thread.extraction_done.connect(self.extraction_complete)

        self.set_ui_enabled(False)
        self.report_text.clear()
        self.progress_bar.setValue(0)
        self.thread.start()

    def update_progress(self, value, time_remaining):
        self.progress_bar.setValue(value)
        self.time_label.setText(f"⏳ {time_remaining}")
        QApplication.processEvents()

    def update_status(self, message):
        self.status_bar.showMessage(message)
        self.time_label.setText(message)
        QApplication.processEvents()

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

            latest_file = data.get('latest_archive', 'N/A')
            latest_file_ctime = data.get('latest_archive_ctime', 'N/A')
            latest_file_mtime = data.get('latest_archive_mtime', 'N/A')

            report_lines.extend([
                f"\n📁 {folder_name}",
                f"   {status_icon} Status: {data['status']}",
                f"   📦 Arquivo mais recente: {latest_file}",
                f"   🕒 Criado em: {latest_file_ctime}",
                f"   🕒 Modificado em: {latest_file_mtime}",
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

    def open_extracted_folder(self, item):
        # Abre a pasta extraída correspondente no Explorer
        folder = item.text()
        latest_file = ExtractionThread().find_latest_archive(folder)
        if latest_file:
            archive_name = os.path.splitext(os.path.basename(latest_file))[0]
            extracted_folder = os.path.join(folder, f"extracted_{archive_name}")
            if os.path.exists(extracted_folder):
                os.startfile(extracted_folder)
            else:
                QMessageBox.warning(self, "Aviso", "Pasta extraída não encontrada.")
        else:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo compactado encontrado.")

    def select_all_folders(self):
        for i in range(self.folder_list.count()):
            self.folder_list.item(i).setCheckState(Qt.Checked)

    def unselect_all_folders(self):
        for i in range(self.folder_list.count()):
            self.folder_list.item(i).setCheckState(Qt.Unchecked)

    def delete_extracted_folders(self):
        selected = [
            self.folder_list.item(i).text()
            for i in range(self.folder_list.count())
            if self.folder_list.item(i).isSelected()
        ]
        if not selected:
            QMessageBox.information(self, "Atenção", "Selecione pelo menos uma pasta na lista para excluir a extração.")
            return
        confirm = QMessageBox.question(
            self, "Confirmar Exclusão",
            f"Tem certeza que deseja excluir as extrações selecionadas?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            for folder in selected:
                latest_file = ExtractionThread().find_latest_archive(folder)
                if latest_file:
                    archive_name = os.path.splitext(os.path.basename(latest_file))[0]
                    extracted_folder = os.path.join(folder, f"extracted_{archive_name}")
                    if os.path.exists(extracted_folder):
                        try:
                            shutil.rmtree(extracted_folder)
                        except Exception as e:
                            QMessageBox.warning(self, "Erro", f"Erro ao excluir {extracted_folder}: {e}")
            QMessageBox.information(self, "Concluído", "Extração(ões) excluída(s) com sucesso.")

if __name__ == "__main__":
    try:
        import rarfile
    except ImportError:
        print("Erro: Instale rarfile com: pip install rarfile")
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    
    # Configurar fonte padrão
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    window = BackupExtractor()
    window.show()
    sys.exit(app.exec_())