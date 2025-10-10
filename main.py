import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QFrame
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt

from data_processing import analyze_csv


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV Analyzer (доставки)")
        self.setMinimumSize(560, 300)

        # Состояние
        self.selected_path: Path | None = None

        # Виджеты
        self.pick_btn = QPushButton("Выбрать CSV…")
        self.run_btn  = QPushButton("Показать результат")
        self.run_btn.setEnabled(False)

        self.file_label = QLabel("Файл не выбран")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #666;")

        # Результаты
        self.total_label = QLabel("Выполнено всего: —")
        self.p_label     = QLabel("Постоматы: —")
        self.other_label = QLabel("Остальные: —")
        for w in (self.total_label, self.p_label, self.other_label):
            w.setStyleSheet("font-size: 14px;")


        # Лэйауты
        top = QHBoxLayout()
        top.addWidget(self.pick_btn)
        top.addWidget(self.run_btn)
        top.addStretch()

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)

        results = QVBoxLayout()
        results.addWidget(self.total_label)
        results.addWidget(self.p_label)
        results.addWidget(self.other_label)

        root = QVBoxLayout()
        root.addLayout(top)
        root.addWidget(self.file_label)
        root.addWidget(sep)
        root.addLayout(results)
        root.addStretch()

        container = QWidget()
        container.setLayout(root)
        self.setCentralWidget(container)

        # Сигналы
        self.pick_btn.clicked.connect(self.pick_file)
        self.run_btn.clicked.connect(self.run_analysis)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите CSV-файл",
            "",
            "CSV Files (*.csv);;All Files (*)",
        )
        if not path:
            return
        self.selected_path = Path(path)
        self.file_label.setText(f"Выбран файл: {self.selected_path}")
        self.run_btn.setEnabled(True)
        # Сброс предыдущих результатов
        self.total_label.setText("Выполнено всего: —")
        self.p_label.setText("Постоматы: —")
        self.other_label.setText("Остальные: —")

    def run_analysis(self):
        if not self.selected_path:
            QMessageBox.information(self, "Файл не выбран", "Сначала выбери CSV-файл.")
            return
        try:
            res = analyze_csv(self.selected_path)
            self.total_label.setText(f"Выполнено всего: {res['total_completed']}")
            self.p_label.setText(f"Постоматы: {res['postomats']}")
            self.other_label.setText(f"Остальные: {res['others']}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка анализа", str(e))


def main():
    app = QApplication(sys.argv)
    # Иконка по желанию:
    # app.setWindowIcon(QIcon("icon.png"))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()