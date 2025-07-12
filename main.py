import os
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, 
                            QLineEdit, QPushButton, QFileDialog, QMessageBox, QTextEdit)
from PyQt5.QtCore import Qt
import zipfile
from datetime import datetime

# Estilo integrado diretamente no arquivo principal
APP_STYLE = """
/* Estilo Dark Moderno */
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

QLineEdit:focus {
    border: 1px solid #555;
}

QTextEdit {
    background-color: #3A3A3A;
    border: 1px solid #444;
    border-radius: 4px;
    padding: 8px;
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

QScrollBar:vertical {
    background: #3A3A3A;
    width: 12px;
}

QScrollBar::handle:vertical {
    background: #505050;
    min-height: 20px;
    border-radius: 6px;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

class ZIPExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Descompactador ZIP Avançado")
        self.setGeometry(200, 200, 800, 600)
        
        self.folder_path = ""
        self.password = ""
        self.report_data = []
        
        self.initUI()
        
    def initUI(self):
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Título
        title = QLabel("Descompactador de Arquivos ZIP")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title, alignment=Qt.AlignCenter)
        
        # Seção de seleção de pasta
        folder_group = QWidget()
        folder_layout = QVBoxLayout()
        
        self.folder_label = QLabel("📁 Pasta selecionada: Nenhuma pasta selecionada")
        folder_layout.addWidget(self.folder_label)
        
        self.select_folder_btn = QPushButton("Selecionar Pasta")
        self.select_folder_btn.setStyleSheet("background-color: #4A6FA5;")
        self.select_folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.select_folder_btn)
        
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)
        
        # Seção de senha
        password_group = QWidget()
        password_layout = QVBoxLayout()
        
        self.password_label = QLabel("🔑 Senha:")
        password_layout.addWidget(self.password_label)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Digite a senha para os arquivos ZIP")
        self.password_input.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(self.password_input)
        
        password_group.setLayout(password_layout)
        layout.addWidget(password_group)
        
        # Botão de extração
        self.extract_btn = QPushButton("▶ Descompactar Arquivos")
        self.extract_btn.setStyleSheet("background-color: #5A8E5A; font-weight: bold;")
        self.extract_btn.clicked.connect(self.extract_files)
        self.extract_btn.setEnabled(False)
        layout.addWidget(self.extract_btn)
        
        # Seção de relatório
        report_group = QWidget()
        report_layout = QVBoxLayout()
        
        self.report_label = QLabel("📊 Relatório:")
        report_layout.addWidget(self.report_label)
        
        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        self.report_text.setStyleSheet("font-family: 'Consolas', monospace;")
        report_layout.addWidget(self.report_text)
        
        report_group.setLayout(report_layout)
        layout.addWidget(report_group, stretch=1)
        
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)
        
        # Barra de status
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Pronto para começar")
    
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"📁 Pasta selecionada: {folder}")
            self.extract_btn.setEnabled(True)
            self.status_bar.showMessage(f"Pasta pronta: {os.path.basename(folder)}", 3000)
    
    def extract_files(self):
        self.password = self.password_input.text()
        if not self.folder_path:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione uma pasta antes de continuar.")
            return
            
        zip_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith('.zip')]
        if not zip_files:
            QMessageBox.information(self, "Informação", "Nenhum arquivo ZIP encontrado na pasta selecionada.")
            return
            
        self.report_data = []
        success_count = 0
        error_count = 0
        
        for zip_file in zip_files:
            file_path = os.path.join(self.folder_path, zip_file)
            output_folder = os.path.join(self.folder_path, os.path.splitext(zip_file)[0])
            
            try:
                # Create output folder if it doesn't exist
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
                
                # Extract files
                with zipfile.ZipFile(file_path) as zf:
                    if self.password:
                        zf.extractall(path=output_folder, pwd=self.password.encode('utf-8'))
                    else:
                        zf.extractall(path=output_folder)
                
                # Get original file size in MB
                file_size = os.path.getsize(file_path) / (1024 * 1024)
                
                self.report_data.append({
                    'file_name': zip_file,
                    'size_mb': round(file_size, 2),
                    'status': '✅ Sucesso',
                    'message': ''
                })
                success_count += 1
                
            except Exception as e:
                error_msg = str(e)
                file_size = os.path.getsize(file_path) / (1024 * 1024) if os.path.exists(file_path) else 0
                
                self.report_data.append({
                    'file_name': zip_file,
                    'size_mb': round(file_size, 2) if file_size else 0,
                    'status': '❌ Erro',
                    'message': error_msg
                })
                error_count += 1
                
                # Show error message for each file
                QMessageBox.critical(self, "Erro", f"Erro ao descompactar {zip_file}:\n{error_msg}")
        
        # Generate report
        self.generate_report()
        
        # Show summary
        QMessageBox.information(self, "Conclusão", 
                               f"Processo concluído!\n\n"
                               f"📦 Arquivos processados: {len(zip_files)}\n"
                               f"✅ Sucessos: {success_count}\n"
                               f"❌ Erros: {error_count}")
    
    def generate_report(self):
        if not self.report_data:
            self.report_text.setPlainText("Nenhum dado para relatório.")
            return
            
        report_lines = []
        report_lines.append(f"📝 Relatório de Descompactação - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"📂 Pasta: {self.folder_path}")
        report_lines.append("\n📋 Detalhes dos Arquivos:")
        report_lines.append("=" * 80)
        report_lines.append(f"{'Arquivo':<40} {'Tamanho (MB)':>12} {'Status':<12} {'Mensagem'}")
        report_lines.append("=" * 80)
        
        total_size = 0
        
        for item in self.report_data:
            report_lines.append(f"{item['file_name']:<40} {item['size_mb']:>12.2f} {item['status']:<12} {item['message']}")
            total_size += item['size_mb']
        
        report_lines.append("=" * 80)
        report_lines.append(f"{'TOTAL:':<40} {total_size:>12.2f} MB")
        
        self.report_text.setPlainText("\n".join(report_lines))
        
        # Save report to file
        report_file = os.path.join(self.folder_path, "relatorio_descompactacao.txt")
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("\n".join(report_lines))
        
        self.status_bar.showMessage(f"Relatório salvo em: {report_file}", 5000)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Aplicar o estilo
    app.setStyleSheet(APP_STYLE)
    
    window = ZIPExtractor()
    window.show()
    sys.exit(app.exec_())