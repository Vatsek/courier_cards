import sys
import re
from pathlib import Path
from typing import List
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit
)
from PyQt6.QtCore import Qt

from data_processing import analyze_csvs, process_kt_excels, analyze_pm_excels


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CDEK helper")
        self.setMinimumSize(860, 520)

        # Состояние
        self.selected_paths: List[Path] = []

        # Кнопки
        self.pick_btn = QPushButton("Выбрать файлы…")
        self.run_btn  = QPushButton("Посчитать доставки")
        self.kt_btn   = QPushButton("КТ (очистить Excel)")
        self.pm_btn   = QPushButton("Показать ПМ")

        self.run_btn.setEnabled(False)
        self.kt_btn.setEnabled(False)
        self.pm_btn.setEnabled(False)

        # Инфо о файлах
        self.file_label = QLabel("Файлы не выбраны")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #ccc;")

        # Блок для вывода результатов
        self.result_box = QVBoxLayout()
        self.result_labels: list[QLabel] = []

        def show_result_lines(lines: list[str]):
            for lbl in self.result_labels:
                lbl.deleteLater()
            self.result_labels.clear()
            for line in lines:
                lbl = QLabel(line)
                lbl.setStyleSheet("font-size: 14px; color: #eee;")
                lbl.setTextInteractionFlags(
                    Qt.TextInteractionFlag.TextSelectableByMouse |
                    Qt.TextInteractionFlag.TextSelectableByKeyboard
                )
                self.result_box.addWidget(lbl)
                self.result_labels.append(lbl)

        self.show_result_lines = show_result_lines

        # Лог — тёмный, аккуратный, ровный
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Лог по файлам…")
        self.log.setMinimumHeight(160)
        self.log.setStyleSheet("""
            QTextEdit {
                font-family: 'Courier New', monospace;
                font-size: 13px;
                color: #ddd;
                background-color: #1e1e1e;
                border: 1px solid #555;
                padding: 6px;
            }
        """)

        # Верхняя панель
        top = QHBoxLayout()
        top.addWidget(self.pick_btn)
        top.addWidget(self.run_btn)
        top.addWidget(self.kt_btn)
        top.addWidget(self.pm_btn)
        top.addStretch()

        # Основной layout
        root = QVBoxLayout()
        root.addLayout(top)
        root.addWidget(self.file_label)
        root.addLayout(self.result_box)
        root.addWidget(self.log)
        root.addStretch()

        container = QWidget()
        container.setLayout(root)
        self.setCentralWidget(container)

        # Сигналы
        self.pick_btn.clicked.connect(self.pick_files)
        self.run_btn.clicked.connect(self.run_analysis_csv)
        self.kt_btn.clicked.connect(self.run_kt_excels)
        self.pm_btn.clicked.connect(self.run_pm_excels)

    def pick_files(self):
        downloads = str(Path.home() / "Downloads")

        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select CSV/Excel files",
            downloads,
            "CSV/Excel Files (*.csv *.xlsx *.xls *.xlsm);;All Files (*)"
        )
        if not paths:
            return
        self.selected_paths = [Path(p) for p in paths]

        # Подсчёт типов
        csv_count = sum(1 for p in self.selected_paths if p.suffix.lower() == ".csv")
        xls_count = sum(1 for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm"))

        self.file_label.setText(
            f"Выбрано файлов: {len(self.selected_paths)} (CSV: {csv_count}, Excel: {xls_count})"
        )
        self.run_btn.setEnabled(csv_count > 0)
        self.kt_btn.setEnabled(xls_count > 0)
        self.pm_btn.setEnabled(xls_count > 0)

        # Очистить результаты и лог
        self.show_result_lines([])
        self.log.clear()

    def _extract_date_from_name(self, filename: str):
        """Извлекает дату из имени файла (09_10_2025, 2025-10-09 и т.п.)."""
        match = re.search(r"(\d{2})[._-](\d{2})[._-](\d{4})", filename)
        if match:
            day, month, year = match.groups()
            try:
                dt = datetime(int(year), int(month), int(day))
                return f"{day}.{month}.{year}", dt
            except ValueError:
                pass
        return "Нет даты", datetime.min

    def run_analysis_csv(self):
        csv_paths = [p for p in self.selected_paths if p.suffix.lower() == ".csv"]
        if not csv_paths:
            return
        try:
            res = analyze_csvs(csv_paths)
            totals = res["totals"]

            # Основной блок
            self.show_result_lines([
                f"Выполнено всего: {totals['total_completed']}",
                f"Д/З: {totals.get('orders', 0) + totals.get('deliveries', 0)}",
                f"Постоматы: {totals['postomats']}",
                f"Передача на ПВЗ: {totals.get('pvz', 0)}",
            ])

            # Лог по каждому файлу
            log_entries = []
            for f in res["per_file"]:
                date_str, dt = self._extract_date_from_name(Path(f['file']).name)
                log_entries.append({
                    "date_str": date_str,
                    "dt": dt,
                    "dz": f['orders'] + f['deliveries'],
                    "postomats": f['postomats'],
                    "pvz": f['pvz'],
                })

            # Сортировка по дате (по возрастанию)
            log_entries.sort(key=lambda x: x["dt"])

            # Табличный вывод
            lines = ["Дата         | Д/З  | Постоматы | ПВЗ",
                     "---------------------------------------"]
            for e in log_entries:
                lines.append(
                    f"{e['date_str']:<12} | {e['dz']:<4} | {e['postomats']:<9} | {e['pvz']:<3}"
                )

            self.log.clear()
            self.log.append("\n".join(lines))

        except Exception as e:
            self.show_result_lines([f"Ошибка анализа: {e}"])

    def run_kt_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            return
        try:
            res = process_kt_excels(xls_paths)
            self.show_result_lines([f"КТ: создано файлов — {len(res['saved'])}.\nСохранено в загрузки"])
        except Exception as e:
            self.show_result_lines([f"Ошибка КТ: {e}"])

    def run_pm_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            return
        try:
            res = analyze_pm_excels(xls_paths)

            def fmt(v):
                if v is None:
                    return "—"
                try:
                    f = float(v)
                    return ("%.4f" % f).rstrip("0").rstrip(".")
                except Exception:
                    return str(v)

            if res["results"]:
                vals = res["results"][0]["values"]
                self.show_result_lines([
                    f"Декабрьская: {fmt(vals.get('Декабрьская'))}",
                    f"Живова: {fmt(vals.get('Живова'))}",
                    f"Мневники: {fmt(vals.get('Мневники'))}",
                    f"Твардовского: {fmt(vals.get('Твардовского'))}",
                ])
            else:
                self.show_result_lines(["ПМ: нет данных."])

        except Exception as e:
            self.show_result_lines([f"Ошибка ПМ: {e}"])


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()