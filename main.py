import sys
from pathlib import Path
from typing import List
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QFrame, QTextEdit
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
        self.pick_btn = QPushButton("Выбрать файлы…")          # CSV и Excel
        self.run_btn  = QPushButton("Посчитать доставки")      # анализ CSV
        self.kt_btn   = QPushButton("КТ (очистить Excel)")     # очистка Excel
        self.pm_btn   = QPushButton("Показать ПМ")             # последняя миля

        self.run_btn.setEnabled(False)
        self.kt_btn.setEnabled(False)
        self.pm_btn.setEnabled(False)

        # Инфо о файлах
        self.file_label = QLabel("Файлы не выбраны")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("color: #666;")

        # --- Универсальный блок результатов (изначально ПУСТО) ---
        self.result_box = QVBoxLayout()
        self.result_labels: list[QLabel] = []

        def show_result_lines(lines: list[str]):
            # удалить предыдущие строки
            for lbl in self.result_labels:
                lbl.deleteLater()
            self.result_labels.clear()
            # добавить новые
            for line in lines:
                lbl = QLabel(line)
                lbl.setStyleSheet("font-size: 14px;")
                lbl.setTextInteractionFlags(
                    Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard
                )
                self.result_box.addWidget(lbl)
                self.result_labels.append(lbl)

        self.show_result_lines = show_result_lines  # сохранить как метод

        hint = QLabel(
            "«Посчитать доставки» — анализ CSV (строки со статусом «Выполнено»). "
            "«КТ» — очищает Excel (.xlsx/.xls/.xlsm): оставляет нужные столбцы и сохраняет <имя>_KT.xlsx. "
            "«Показать ПМ» — вытягивает метрику «Ср. срок на последней миле для 2 якоря без СДД, дн» "
            "для: Декабрьская=MSK650, Живова=MSK963, Мневники=MSK1125, Твардовского=MSK2469."
        )
        hint.setStyleSheet("color: #888; font-size: 12px;")
        hint.setWordWrap(True)

        # Журнал снизу (можно скрыть, если не нужен)
        self.details = QTextEdit()
        self.details.setReadOnly(True)
        self.details.setPlaceholderText("Журнал операций (опционально).")
        self.details.setMinimumHeight(160)

        # Верхняя панель
        top = QHBoxLayout()
        top.addWidget(self.pick_btn)
        top.addWidget(self.run_btn)
        top.addWidget(self.kt_btn)
        top.addWidget(self.pm_btn)
        top.addStretch()

        # Разделитель
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)

        # Корневой layout
        root = QVBoxLayout()
        root.addLayout(top)
        root.addWidget(self.file_label)
        root.addWidget(sep)
        root.addLayout(self.result_box)   # <-- тут изначально пусто
        root.addWidget(hint)
        root.addWidget(self.details)
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
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите файлы CSV/Excel", "", "CSV/Excel Files (*.csv *.xlsx *.xls *.xlsm);;All Files (*)"
        )
        if not paths:
            return
        self.selected_paths = [Path(p) for p in paths]

        # Подсчёт типов
        csv_count = sum(1 for p in self.selected_paths if p.suffix.lower() == ".csv")
        xls_count = sum(1 for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm"))

        self.file_label.setText(f"Выбрано файлов: {len(self.selected_paths)} (CSV: {csv_count}, Excel: {xls_count})")
        self.run_btn.setEnabled(csv_count > 0)
        self.kt_btn.setEnabled(xls_count > 0)
        self.pm_btn.setEnabled(xls_count > 0)

        # Очистить основной блок результатов и журнал при новом выборе файлов
        self.show_result_lines([])
        self.details.clear()

    def run_analysis_csv(self):
        csv_paths = [p for p in self.selected_paths if p.suffix.lower() == ".csv"]
        if not csv_paths:
            QMessageBox.information(self, "Нет CSV", "Выбери хотя бы один CSV-файл.")
            return
        try:
            res = analyze_csvs(csv_paths)
            totals = res["totals"]

            # Выводим только после нажатия
            self.show_result_lines([
                f"Выполнено всего: {totals['total_completed']}",
                f"Постоматы: {totals['postomats']}",
                f"Доставки/Заявки: {totals['others']}",
            ])

            # Подробности — в журнал (опционально)
            lines = []
            if res["per_file"]:
                lines.append("CSV — детали по файлам:")
                for row in res["per_file"]:
                    lines.append(
                        f"- {Path(row['file']).name}: выполнено={row['total_completed']}, "
                        f"Постоматы={row['postomats']}, Доставки/Заявки={row['others']}"
                    )
            if res["errors"]:
                lines.append("\nCSV — ошибки:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")
            if lines:
                self.details.append("\n".join(lines))
        except Exception as e:
            QMessageBox.critical(self, "Ошибка анализа CSV", str(e))

    def run_kt_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            QMessageBox.information(self, "Нет Excel", "Выбери хотя бы один Excel-файл.")
            return
        try:
            res = process_kt_excels(xls_paths)

            # Короткий итог — только по кнопке
            self.show_result_lines([f"КТ: создано файлов — {len(res['saved'])}"])

            # Детали в журнал
            lines = []
            if res["saved"]:
                lines.append("КТ — сохранённые файлы:")
                for r in res["saved"]:
                    lines.append(f"- {Path(r['file']).name} → {Path(r['saved_as']).name}")
            if res["skipped"]:
                lines.append("\nКТ — пропущены (не Excel):")
                for s in res["skipped"]:
                    lines.append(f"- {Path(s['file']).name}: {s['reason']}")
            if res["errors"]:
                lines.append("\nКТ — ошибки:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")
            if lines:
                self.details.append("\n".join(lines))

            # QMessageBox.information(self, "КТ завершено", "Обработка Excel-файлов завершена.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка КТ", str(e))

    def run_pm_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            QMessageBox.information(self, "Нет Excel", "Выбери хотя бы один Excel-файл.")
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

            # Показываем только после нажатия
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

            # В журнал — все результаты
            lines = []
            for item in res["results"]:
                vals = item["values"]
                lines.append(f"ПМ — файл: {Path(item['file']).name} (лист: {item.get('sheet', '—')})")
                lines.append(f"  Декабрьская: {fmt(vals.get('Декабрьская'))}")
                lines.append(f"  Живова:      {fmt(vals.get('Живова'))}")
                lines.append(f"  Мневники:    {fmt(vals.get('Мневники'))}")
                lines.append(f"  Твардовского:{fmt(vals.get('Твардовского'))}")
                lines.append("")
            if res["errors"]:
                lines.append("ПМ — ошибки:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")
            if res["skipped"]:
                lines.append("\nПМ — пропущены (не Excel):")
                for s in res["skipped"]:
                    lines.append(f"- {Path(s['file']).name}: {s['reason']}")
            if lines:
                self.details.append("\n".join(lines).strip())

            # QMessageBox.information(self, "ПМ готово", "Значения последней мили извлечены.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка ПМ", str(e))


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()