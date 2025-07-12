import os
import sys
import zipfile
import rarfile
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, 
                           QLineEdit, QPushButton, QFileDialog, QMessageBox, QTextEdit,
                           QHBoxLayout, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

APP_STYLE = """
QWidget {
    background-color: #2D2D2D;
    color: #E0E0E0;
    font-family: 'Segoe UI';
    font-size: 12px;
    border: none;
}
QMainWindow {
    background-color: #2D2D2D;
    border: 1px solid #444;
}
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
QLineEdit {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
    selection-background-color: #505050;
}
QTextEdit {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
    font-family: 'Consolas', monospace;
}
QLabel {
    color: #E0E0E0;
    padding: 4px 0;
}
QStatusBar {
    background-color: #3A3A3A;
    border-top: 1px solid #444;
    padding: 4px;
}
QProgressBar {
    border: 1px solid #444;
    border-radius: 4px;
    text-align: center;
    background-color: #3A3A3A;
}
QProgressBar::chunk {
    background-color: #4A6FA5;
    width: 10px;
}
"""

class ExtractionThread(QThread):
    update_progress = pyqtSignal(int)
    update_status = pyqtSignal(str)
    extraction_done = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.root_folder = ""
        self.password = ""

    def find_archive_folders(self):
        archive_folders = []
        for root, dirs, files in os.walk(self.root_folder):
            if any(f.lower().endswith(('.zip', '.rar')) for f in files):
                archive_folders.append(root)
        return archive_folders

    def find_latest_archive(self, folder):
        valid_files = []
        for f in os.listdir(folder):
            if f.lower().endswith(('.zip', '.rar')):
                file_path = os.path.join(folder, f)
                mtime = os.path.getmtime(file_path)
                valid_files.append((file_path, mtime))
        return max(valid_files, key=lambda x: x[1])[0] if valid_files else None

    def extract_archive(self, archive_path, output_base):
        try:
            archive_name = os.path.basename(archive_path)
            output_folder = os.path.join(output_base, "extracted_" + os.path.splitext(archive_name)[0])
            os.makedirs(output_folder, exist_ok=True)
            original_size = os.path.getsize(archive_path) / (1024 * 1024)  # Tamanho original em MB

            if archive_path.lower().endswith('.zip'):
                with zipfile.ZipFile(archive_path) as zf:
                    zf.extractall(path=output_folder, pwd=self.password.encode('utf-8') if self.password else None)
            else:  # RAR
                with rarfile.RarFile(archive_path) as rf:
                    rf.extractall(path=output_folder, pwd=self.password if self.password else None)
            
            # Calcular tamanho total dos arquivos extraídos
            extracted_size = 0
            extracted_files = []
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    extracted_size += os.path.getsize(file_path)
                    extracted_files.append(file)
            
            extracted_size_mb = extracted_size / (1024 * 1024)  # Converter para MB

            return {
                "status": "Sucesso",
                "message": f"Extraídos {len(extracted_files)} arquivos",
                "original_size_mb": round(original_size, 2),
                "extracted_size_mb": round(extracted_size_mb, 2),
                "files": extracted_files
            }
        except Exception as e:
            return {
                "status": "Erro",
                "message": str(e),
                "original_size_mb": 0,
                "extracted_size_mb": 0,
                "files": []
            }

    def run(self):
        archive_folders = self.find_archive_folders()
        total_results = {}
        
        if not archive_folders:
            self.update_status.emit("Nenhum arquivo ZIP/RAR encontrado")
            self.extraction_done.emit({})
            return
            
        for i, folder in enumerate(archive_folders):
            self.update_status.emit(f"Processando: {os.path.basename(folder)}...")
            latest_file = self.find_latest_archive(folder)
            if not latest_file:
                total_results[folder] = {"status": "Ignorado", "message": "Nenhum arquivo válido"}
                continue
            
            result = self.extract_archive(latest_file, folder)
            total_results[folder] = result
            self.update_progress.emit(int((i + 1) / len(archive_folders) * 100))
        
        self.extraction_done.emit(total_results)

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

    def update_progress(self, value):
        self.progress_bar.setValue(value)

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
                f"   💬 Mensagem: {data['message']}"
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