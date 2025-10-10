import sys
from pathlib import Path
from typing import List
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QFrame, QTextEdit
)
from PyQt6.QtCore import Qt

from data_processing import analyze_csvs


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV Analyzer (Курьерские карты)")
        self.setMinimumSize(640, 360)

        # Состояние
        self.selected_paths: List[Path] = []

        # Кнопки
        self.pick_btn = QPushButton("Выбрать CSV-файлы…")
        self.run_btn  = QPushButton("Показать результат")
        self.run_btn.setEnabled(False)

        # Инфо о файлах
        self.file_label = QLabel("Файлы не выбраны")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #666;")

        # Результаты (итог по всем файлам)
        self.total_label = QLabel("Выполнено всего: —")
        self.p_label     = QLabel("Постоматы: —")
        self.other_label = QLabel("Доставки/Заявки: —")
        for w in (self.total_label, self.p_label, self.other_label):
            w.setStyleSheet("font-size: 14px;")

        # Подробности (перечень файлов и возможные ошибки)
        self.details = QTextEdit()
        self.details.setReadOnly(True)
        self.details.setPlaceholderText("Здесь появятся детали по каждому файлу и сообщения об ошибках (если будут).")
        self.details.setMinimumHeight(120)

        hint = QLabel("Подсказка: учитываются только строки со статусом «Выполнено».")
        hint.setStyleSheet("color: #888; font-size: 12px;")
        hint.setWordWrap(True)

        # Лэйаут: верхняя панель
        top = QHBoxLayout()
        top.addWidget(self.pick_btn)
        top.addWidget(self.run_btn)
        top.addStretch()

        # Разделитель
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)

        # Итоговый блок
        totals_box = QVBoxLayout()
        totals_box.addWidget(self.total_label)
        totals_box.addWidget(self.p_label)
        totals_box.addWidget(self.other_label)

        # Корневой лэйаут
        root = QVBoxLayout()
        root.addLayout(top)
        root.addWidget(self.file_label)
        root.addWidget(sep)
        root.addLayout(totals_box)
        root.addWidget(hint)
        root.addWidget(self.details)
        root.addStretch()

        container = QWidget()
        container.setLayout(root)
        self.setCentralWidget(container)

        # Сигналы
        self.pick_btn.clicked.connect(self.pick_files)
        self.run_btn.clicked.connect(self.run_analysis)

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите один или несколько CSV-файлов",
            "",
            "CSV Files (*.csv);;All Files (*)",
        )
        if not paths:
            return
        self.selected_paths = [Path(p) for p in paths]
        self.file_label.setText(f"Выбрано файлов: {len(self.selected_paths)}")
        self.run_btn.setEnabled(True)

        # Сброс результата
        self.total_label.setText("Выполнено всего: —")
        self.p_label.setText("Постоматы: —")
        self.other_label.setText("Доставки/Заявки: —")
        self.details.clear()

    def run_analysis(self):
        if not self.selected_paths:
            QMessageBox.information(self, "Файлы не выбраны", "Сначала выбери один или несколько CSV-файлов.")
            return
        try:
            res = analyze_csvs(self.selected_paths)
            totals = res["totals"]
            self.total_label.setText(f"Выполнено всего: {totals['total_completed']}")
            self.p_label.setText(f"Постоматы: {totals['postomats']}")
            self.other_label.setText(f"Доставки/Заявки: {totals['others']}")

            # Детали: пер-файл + ошибки
            lines = []
            if res["per_file"]:
                lines.append("Детали по файлам:")
                for row in res["per_file"]:
                    lines.append(f"- {Path(row['file']).name}: выполнено={row['total_completed']}, П={row['postomats']}, Доставки/Заявки={row['others']}")
            if res["errors"]:
                lines.append("\n⚠️ Ошибки при обработке:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")
            self.details.setPlainText("\n".join(lines) if lines else "Нет подробностей.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка анализа", str(e))


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()