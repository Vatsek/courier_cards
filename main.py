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
        self.setMinimumSize(820, 480)

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

        # Итоги по CSV
        self.total_label = QLabel("Выполнено всего: —")
        self.post_label  = QLabel("Постоматы: —")
        self.other_label = QLabel("Доставки/Заявки: —")
        for w in (self.total_label, self.post_label, self.other_label):
            w.setStyleSheet("font-size: 14px;")

        hint = QLabel(
            "«Посчитать доставки» считает выбранные CSV (строки со статусом «Выполнено»). "
            "«КТ» очищает выбранные Excel (.xlsx/.xls/.xlsm): оставляет нужные столбцы и сохраняет <имя>_KT.xlsx. "
            "«Показать ПМ» — извлекает метрику «Ср. срок на последней миле для 2 якоря без СДД, дн» "
            "для: Декабрьская=MSK650, Живова=MSK963, Мневники=MSK1125, Твардовского=MSK2469."
        )
        hint.setStyleSheet("color: #888; font-size: 12px;")
        hint.setWordWrap(True)

        # Детали/лог
        self.details = QTextEdit()
        self.details.setReadOnly(True)
        self.details.setPlaceholderText("Здесь появятся детали по анализу CSV, КТ и ПМ.")
        self.details.setMinimumHeight(200)

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

        # Итоги CSV
        totals_box = QVBoxLayout()
        totals_box.addWidget(self.total_label)
        totals_box.addWidget(self.post_label)
        totals_box.addWidget(self.other_label)

        # Root
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
        self.run_btn.clicked.connect(self.run_analysis_csv)
        self.kt_btn.clicked.connect(self.run_kt_excels)
        self.pm_btn.clicked.connect(self.run_pm_excels)

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы CSV/Excel",
            "",
            "CSV/Excel Files (*.csv *.xlsx *.xls *.xlsm);;All Files (*)",
        )
        if not paths:
            return
        self.selected_paths = [Path(p) for p in paths]

        # Подсчёт типов для статуса
        csv_count = sum(1 for p in self.selected_paths if p.suffix.lower() == ".csv")
        xls_count = sum(1 for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm"))

        self.file_label.setText(f"Выбрано файлов: {len(self.selected_paths)} (CSV: {csv_count}, Excel: {xls_count})")
        self.run_btn.setEnabled(csv_count > 0)
        self.kt_btn.setEnabled(xls_count > 0)
        self.pm_btn.setEnabled(xls_count > 0)

        # Сброс результатов CSV
        self.total_label.setText("Выполнено всего: —")
        self.post_label.setText("Постоматы: —")
        self.other_label.setText("Доставки/Заявки: —")

        # Очистка лога
        self.details.clear()

    def run_analysis_csv(self):
        csv_paths = [p for p in self.selected_paths if p.suffix.lower() == ".csv"]
        if not csv_paths:
            QMessageBox.information(self, "Нет CSV", "Выбери хотя бы один CSV-файл.")
            return
        try:
            res = analyze_csvs(csv_paths)
            totals = res["totals"]
            self.total_label.setText(f"Выполнено всего: {totals['total_completed']}")
            self.post_label.setText(f"Постоматы: {totals['postomats']}")
            self.other_label.setText(f"Доставки/Заявки: {totals['others']}")

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
            self.details.append("\n".join(lines) if lines else "CSV: нет подробностей.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка анализа CSV", str(e))

    def run_kt_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            QMessageBox.information(self, "Нет Excel", "Выбери хотя бы один Excel-файл (.xlsx/.xls/.xlsm).")
            return
        try:
            res = process_kt_excels(xls_paths)
            lines = []

            if res["saved"]:
                lines.append("КТ — сохранённые файлы:")
                for r in res["saved"]:
                    src = Path(r["file"]).name
                    dst = Path(r["saved_as"]).name
                    kept = ", ".join(r["kept_columns"])
                    lines.append(f"- {src} → {dst} (оставлены столбцы: {kept})")

            if res["skipped"]:
                lines.append("\nКТ — пропущены (не Excel):")
                for s in res["skipped"]:
                    lines.append(f"- {Path(s['file']).name}: {s['reason']}")

            if res["errors"]:
                lines.append("\nКТ — ошибки:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")

            self.details.append("\n".join(lines) if lines else "КТ: нет изменений.")
            QMessageBox.information(self, "КТ завершено", "Обработка Excel-файлов завершена. Смотри детали внизу.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка КТ", str(e))

    def run_pm_excels(self):
        xls_paths = [p for p in self.selected_paths if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")]
        if not xls_paths:
            QMessageBox.information(self, "Нет Excel", "Выбери хотя бы один Excel-файл (.xlsx/.xls/.xlsm).")
            return
        try:
            res = analyze_pm_excels(xls_paths)

            lines = []
            for item in res["results"]:
                vals = item["values"]

                def fmt(v):
                    if v is None:
                        return "—"
                    try:
                        f = float(v)
                        s = ("%.4f" % f).rstrip("0").rstrip(".")
                        return s
                    except Exception:
                        return str(v)

                lines.append(f"ПМ — файл: {Path(item['file']).name} (лист: {item.get('sheet', '—')})")
                lines.append(f"  Декабрьская: {fmt(vals.get('Декабрьская'))}")
                lines.append(f"  Живова:      {fmt(vals.get('Живова'))}")
                lines.append(f"  Мневники:    {fmt(vals.get('Мневники'))}")
                lines.append(f"  Твардовского:{fmt(vals.get('Твардовского'))}")
                lines.append("")  # пустая строка между файлами

            if res["errors"]:
                lines.append("ПМ — ошибки:")
                for err in res["errors"]:
                    lines.append(f"- {Path(err['file']).name}: {err['error']}")

            if res["skipped"]:
                lines.append("\nПМ — пропущены (не Excel):")
                for s in res["skipped"]:
                    lines.append(f"- {Path(s['file']).name}: {s['reason']}")

            text = "\n".join(lines).strip() or "ПМ: нет данных."
            self.details.append(text)

            QMessageBox.information(self, "ПМ готово", "Значения последней мили извлечены. Смотри детали внизу.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка ПМ", str(e))


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()