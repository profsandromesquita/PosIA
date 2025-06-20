import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QFileDialog, QTextEdit, QMessageBox, QScrollArea
)
from PyQt5.QtGui import QFont, QIcon, QPixmap, QCursor, QImage
from PyQt5.QtCore import Qt
import pandas as pd
import docx2txt
import fitz  # PyMuPDF para PDF

class PyTravelAnalyzer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyTravel - Análise de Destinos")
        self.setGeometry(100, 100, 900, 700)
        self.setStyleSheet("background-color: #002244;")

        self.file_path = ""
        self.catalog_text = ""
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Cabeçalho com logo centralizado e título
        header_container = QWidget()
        header_inner_layout = QVBoxLayout(header_container)

        logo_label = QLabel()
        logo_pixmap = QPixmap("logo.png").scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(logo_pixmap)
        logo_label.setAlignment(Qt.AlignCenter)

        title_label = QLabel("Analisador de Catálogo de Destinos")
        title_label.setFont(QFont("Arial", 20))
        title_label.setStyleSheet("color: yellow;")
        title_label.setAlignment(Qt.AlignCenter)

        header_inner_layout.addWidget(logo_label)
        header_inner_layout.addWidget(title_label)
        layout.addWidget(header_container)

        # Botão de seleção de arquivo
        self.upload_btn = QPushButton("Selecionar Arquivo")
        self.upload_btn.clicked.connect(self.load_file)
        self.upload_btn.setStyleSheet(self.button_style(large=True))
        self.upload_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.upload_btn.setFixedWidth(325)
        layout.addWidget(self.upload_btn, alignment=Qt.AlignHCenter)

        # Nome do arquivo
        self.filename_label = QLabel("")
        self.filename_label.setStyleSheet("color: white; font-size: 14px;")
        self.filename_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.filename_label)

        # Visualizador de PDF renderizado como imagem
        self.pdf_scroll_area = QScrollArea()
        self.pdf_scroll_area.setVisible(False)
        self.pdf_container = QWidget()
        self.pdf_layout = QVBoxLayout(self.pdf_container)
        self.pdf_scroll_area.setWidgetResizable(True)
        self.pdf_scroll_area.setWidget(self.pdf_container)
        layout.addWidget(self.pdf_scroll_area)

        # Área do texto extraído
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setFixedHeight(250)
        self.text_area.setStyleSheet("background-color: #DCDCDC; border: 2px dashed #002244; color: #5A6E7F; font-size: 16px;")
        self.text_area.setText("Clique no botão Selecionar Arquivo para adicionar o documento com os pacotes de viagens disponíveis para que o PyTravel APP possa analisá-lo e sugerir novos pacotes de viagens que irão impulsionar seu negócio.")
        layout.addWidget(self.text_area)

        # Botão confirmar e analisar
        self.analyze_btn = QPushButton("Confirmar e Analisar")
        self.analyze_btn.clicked.connect(self.analyze_file)
        self.analyze_btn.setStyleSheet(self.button_style(large=True))
        self.analyze_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.analyze_btn.setEnabled(False)
        layout.addWidget(self.analyze_btn)

        self.setLayout(layout)

    def button_style(self, large=False):
        size = "padding: 15px; font-size: 20px;" if large else "padding: 10px; font-size: 16px;"
        return (
            f"QPushButton {{ background-color: yellow; color: black; border: none; border-radius: 15px; {size} }}"
            "QPushButton:hover { background-color: #FFEA00; color: #002244; }"
        )

    def load_file(self):
        options = QFileDialog.Options()
        path, _ = QFileDialog.getOpenFileName(self, "Selecione o Catálogo", "", "Documentos (*.docx *.pdf)", options=options)
        if path:
            self.file_path = path
            filename = os.path.basename(path)
            self.filename_label.setText(f"Arquivo carregado: {filename}")
            self.catalog_text = self.read_file(path)
            self.text_area.setText(self.catalog_text)
            self.analyze_btn.setEnabled(True)

            if path.endswith(".pdf"):
                self.show_pdf_as_images(path)
                self.pdf_scroll_area.setVisible(True)
            else:
                self.pdf_scroll_area.setVisible(False)

    def show_pdf_as_images(self, path):
        for i in reversed(range(self.pdf_layout.count())):
            widget_to_remove = self.pdf_layout.itemAt(i).widget()
            if widget_to_remove:
                widget_to_remove.setParent(None)

        doc = fitz.open(path)
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            label = QLabel()
            label.setPixmap(QPixmap.fromImage(image))
            label.setAlignment(Qt.AlignCenter)
            self.pdf_layout.addWidget(label)

    def read_file(self, path):
        if path.endswith(".docx"):
            return docx2txt.process(path)
        elif path.endswith(".pdf"):
            with fitz.open(path) as doc:
                return "\n".join([page.get_text() for page in doc])
        return ""

    def analyze_file(self):
        box = QMessageBox(self)
        box.setWindowTitle("Confirmar Análise")
        box.setText("<span style='color:white;'>Deseja realizar a análise deste catálogo?</span>")
        box.setIcon(QMessageBox.Question)

        yes_button = box.addButton("Yes", QMessageBox.YesRole)
        no_button = box.addButton("No", QMessageBox.NoRole)

        yes_button.setFixedSize(80, 30)
        no_button.setFixedSize(80, 30)

        yes_button.setStyleSheet("background-color: yellow; color: #002244; border: 1px solid #6699CC;")
        no_button.setStyleSheet("background-color: yellow; color: #002244; border: 1px solid #6699CC;")

        box.exec_()

        if box.clickedButton() == yes_button:
            dados = self.processar_catalogo(self.catalog_text)

            options = QFileDialog.Options()
            save_path, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo Excel", "", "CSV Files (*.csv)",
                                                       options=options)
            if save_path:
                df = pd.DataFrame(dict([(k.upper(), pd.Series(v)) for k, v in dados.items()]))
                df.to_csv(save_path, sep=';', index=False, encoding='utf-8-sig')

                confirm = QMessageBox(self)
                confirm.setWindowTitle("Sucesso")
                confirm.setText("<span style='color:white;'>Arquivo salvo com sucesso!</span>")
                confirm.setIcon(QMessageBox.Information)
                ok_button = confirm.addButton("OK", QMessageBox.AcceptRole)
                ok_button.setFixedSize(80, 30)
                ok_button.setStyleSheet("background-color: yellow; color: #002244; border: 1px solid #6699CC;")
                confirm.exec_()

                nova = QMessageBox(self)
                nova.setWindowTitle("Nova Análise")
                nova.setText("<span style='color:white;'>Deseja realizar outra análise?</span>")
                nova.setIcon(QMessageBox.Question)
                sim_button = nova.addButton("Yes", QMessageBox.YesRole)
                nao_button = nova.addButton("No", QMessageBox.NoRole)

                sim_button.setFixedSize(80, 30)
                nao_button.setFixedSize(80, 30)

                sim_button.setStyleSheet("background-color: yellow; color: #002244; border: 1px solid #6699CC;")
                nao_button.setStyleSheet("background-color: yellow; color: #002244; border: 1px solid #6699CC;")

                nova.exec_()

                if nova.clickedButton() == sim_button:
                    self.reset_app()
                else:
                    self.close()

    def processar_catalogo(self, texto):
        secoes = {
            "praias": [], "capitais": [], "interior": [],
            "aviao": [], "onibus": [], "navio": []
        }
        atual = None
        for linha in texto.split('\n'):
            linha = linha.strip()
            if "praianas" in linha.lower():
                atual = "praias"
            elif "capitais" in linha.lower():
                atual = "capitais"
            elif "interior" in linha.lower():
                atual = "interior"
            elif "avião" in linha.lower() or "aviao" in linha.lower():
                atual = "aviao"
            elif "ônibus" in linha.lower() or "onibus" in linha.lower():
                atual = "onibus"
            elif "navio" in linha.lower():
                atual = "navio"
            elif atual and '(' in linha:
                secoes[atual].append(linha)

        cruzamentos = {
            "capitais_praianas": list(set(secoes["capitais"]) & set(secoes["praias"])),
            "praias_onibus": list(set(secoes["praias"]) & set(secoes["onibus"])),
            "interior_aviao": list(set(secoes["interior"]) & set(secoes["aviao"])),
            "capitais_navio": list(set(secoes["capitais"]) & set(secoes["navio"]))
        }
        secoes.update(cruzamentos)
        return secoes

    def reset_app(self):
        self.file_path = ""
        self.catalog_text = ""
        self.filename_label.setText("")
        self.text_area.setText("Clique no botão Selecionar Arquivo para adicionar o documento com os pacotes de viagens disponíveis para que o PyTravel APP possa analisá-lo e sugerir novos pacotes de viagens que irão impulsionar seu negócio.")
        self.analyze_btn.setEnabled(False)
        self.pdf_scroll_area.setVisible(False)
        for i in reversed(range(self.pdf_layout.count())):
            widget_to_remove = self.pdf_layout.itemAt(i).widget()
            if widget_to_remove:
                widget_to_remove.setParent(None)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PyTravelAnalyzer()
    window.show()
    sys.exit(app.exec_())
