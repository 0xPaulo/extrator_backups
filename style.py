# Arquivo de estilos para a aplicação de descompactação

def get_dark_style():
    return """
    /* Estilo geral da aplicação */
    QWidget {
        background-color: #2b2b2b;
        color: #e0e0e0;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 12px;
    }
    
    /* Barra de título */
    QMainWindow {
        background-color: #2b2b2b;
        border: 1px solid #444;
    }
    
    /* Botões */
    QPushButton {
        background-color: #3a3a3a;
        border: 1px solid #444;
        border-radius: 4px;
        padding: 5px 10px;
        min-width: 80px;
    }
    
    QPushButton:hover {
        background-color: #4a4a4a;
        border: 1px solid #555;
    }
    
    QPushButton:pressed {
        background-color: #2a2a2a;
    }
    
    QPushButton:disabled {
        background-color: #2a2a2a;
        color: #777;
    }
    
    /* Campos de texto */
    QLineEdit {
        background-color: #3a3a3a;
        border: 1px solid #444;
        border-radius: 3px;
        padding: 5px;
        selection-background-color: #505050;
    }
    
    QLineEdit:focus {
        border: 1px solid #555;
    }
    
    /* Área de texto */
    QTextEdit {
        background-color: #3a3a3a;
        border: 1px solid #444;
        border-radius: 3px;
        padding: 5px;
    }
    
    /* Labels */
    QLabel {
        color: #e0e0e0;
    }
    
    /* Barra de status */
    QStatusBar {
        background-color: #3a3a3a;
        border-top: 1px solid #444;
    }
    
    /* Caixas de mensagem */
    QMessageBox {
        background-color: #2b2b2b;
    }
    
    QMessageBox QLabel {
        color: #e0e0e0;
    }
    
    QMessageBox QPushButton {
        min-width: 80px;
    }
    
    /* Barra de rolagem */
    QScrollBar:vertical {
        border: none;
        background: #3a3a3a;
        width: 10px;
        margin: 0px 0px 0px 0px;
    }
    
    QScrollBar::handle:vertical {
        background: #505050;
        min-height: 20px;
        border-radius: 4px;
    }
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
    }
    
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
        background: none;
    }
    """

def get_light_style():
    return """
    /* Estilo geral da aplicação */
    QWidget {
        background-color: #f5f5f5;
        color: #333333;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 12px;
    }
    
    /* Barra de título */
    QMainWindow {
        background-color: #f5f5f5;
        border: 1px solid #ddd;
    }
    
    /* Botões */
    QPushButton {
        background-color: #e0e0e0;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 5px 10px;
        min-width: 80px;
        color: #333;
    }
    
    QPushButton:hover {
        background-color: #d0d0d0;
        border: 1px solid #bbb;
    }
    
    QPushButton:pressed {
        background-color: #c0c0c0;
    }
    
    QPushButton:disabled {
        background-color: #e8e8e8;
        color: #999;
    }
    
    /* Campos de texto */
    QLineEdit {
        background-color: #ffffff;
        border: 1px solid #ccc;
        border-radius: 3px;
        padding: 5px;
        selection-background-color: #a0a0a0;
    }
    
    QLineEdit:focus {
        border: 1px solid #aaa;
    }
    
    /* Área de texto */
    QTextEdit {
        background-color: #ffffff;
        border: 1px solid #ccc;
        border-radius: 3px;
        padding: 5px;
    }
    
    /* Labels */
    QLabel {
        color: #333333;
    }
    
    /* Barra de status */
    QStatusBar {
        background-color: #e0e0e0;
        border-top: 1px solid #ccc;
    }
    
    /* Caixas de mensagem */
    QMessageBox {
        background-color: #f5f5f5;
    }
    
    QMessageBox QLabel {
        color: #333333;
    }
    
    QMessageBox QPushButton {
        min-width: 80px;
    }
    
    /* Barra de rolagem */
    QScrollBar:vertical {
        border: none;
        background: #e0e0e0;
        width: 10px;
        margin: 0px 0px 0px 0px;
    }
    
    QScrollBar::handle:vertical {
        background: #c0c0c0;
        min-height: 20px;
        border-radius: 4px;
    }
    
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0px;
    }
    
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
        background: none;
    }
    """