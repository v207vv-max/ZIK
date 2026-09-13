from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.excel.filter import ResultFilter
from app.excel.reader import ExcelReader, ExcelRow
from app.excel.wps_sync import WPSSync
from app.settings import Settings


class MainWindow:
    def __init__(self, controller=None) -> None:
        self.controller = controller

        self.root = tk.Tk()

        self.root.title("ZIK — Smart Call Center")
        self.root.geometry("1450x850")
        self.root.minsize(1150, 700)

        self.settings = Settings()
        self.wps = WPSSync()

        self.reader: ExcelReader | None = None
        self.result_filter: ResultFilter | None = None

        # Все строки из Excel.
        self.all_rows: list[ExcelRow] = []

        # Текущие строки после фильтрации.
        self.rows: list[ExcelRow] = []

        # Индекс текущей строки в self.rows.
        self.current_index = -1

        # Состояние UI.
        self.paused = False
        self.search_text = ""

        self._build_ui()

        if self.controller is not None:
            self.controller.set_status_callback(
                self._on_controller_status
            )
            self.controller.set_log_callback(
                self._on_controller_log
            )

        self._restore_last_file()

    # ============================================================
    # UI
    # ============================================================

    def _build_ui(self) -> None:
        self.root.columnconfigure(
            0,
            weight=1,
        )

        # Основные растягиваемые области.
        self.root.rowconfigure(
            3,
            weight=1,
        )

        self.root.rowconfigure(
            5,
            weight=1,
        )

        # --------------------------------------------------------
        # HEADER
        # --------------------------------------------------------

        header = ttk.Frame(
            self.root,
            padding=(20, 12),
        )

        header.grid(
            row=0,
            column=0,
            sticky="ew",
        )

        header.columnconfigure(
            1,
            weight=1,
        )

        ttk.Label(
            header,
            text="ZIK",
            font=("Segoe UI", 25, "bold"),
        ).grid(
            row=0,
            column=0,
            padx=(0, 20),
        )

        self.status_label = ttk.Label(
            header,
            text="● READY",
            font=("Segoe UI", 11, "bold"),
        )

        self.status_label.grid(
            row=0,
            column=1,
            sticky="w",
        )

        self.counter_label = ttk.Label(
            header,
            text="0 / 0",
            font=("Segoe UI", 11),
        )

        self.counter_label.grid(
            row=0,
            column=2,
        )

        # --------------------------------------------------------
        # EXCEL FILE
        # --------------------------------------------------------

        file_frame = ttk.LabelFrame(
            self.root,
            text=" Excel ",
            padding=10,
        )

        file_frame.grid(
            row=1,
            column=0,
            padx=15,
            pady=(0, 5),
            sticky="ew",
        )

        file_frame.columnconfigure(
            1,
            weight=1,
        )

        ttk.Label(
            file_frame,
            text="Файл:",
        ).grid(
            row=0,
            column=0,
            padx=5,
        )

        self.file_label = ttk.Label(
            file_frame,
            text="Файл не выбран",
        )

        self.file_label.grid(
            row=0,
            column=1,
            sticky="w",
            padx=10,
        )

        ttk.Button(
            file_frame,
            text="📁 Выбрать Excel",
            command=self.select_file,
        ).grid(
            row=0,
            column=2,
            padx=5,
        )

        # --------------------------------------------------------
        # SEARCH / RESULT FILTER
        # --------------------------------------------------------

        search_frame = ttk.Frame(
            self.root,
            padding=(15, 5, 15, 8),
        )

        search_frame.grid(
            row=2,
            column=0,
            sticky="ew",
        )

        ttk.Label(
            search_frame,
            text="🔎 Поиск:",
        ).pack(
            side="left",
            padx=(0, 6),
        )

        self.search_var = tk.StringVar()

        self.search_entry = ttk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=35,
        )

        self.search_entry.pack(
            side="left",
        )

        self.search_entry.bind(
            "<KeyRelease>",
            self._on_search,
        )

        ttk.Button(
            search_frame,
            text="✕",
            width=3,
            command=self.clear_search,
        ).pack(
            side="left",
            padx=5,
        )

        self.result_filter_button = ttk.Button(
            search_frame,
            text="Результат ▾",
            command=self.open_result_filter,
        )

        self.result_filter_button.pack(
            side="left",
            padx=(15, 5),
        )

        ttk.Button(
            search_frame,
            text="Сбросить",
            command=self.clear_all_filters,
        ).pack(
            side="left",
        )

        self.filter_status = ttk.Label(
            search_frame,
            text="Показано: 0",
        )

        self.filter_status.pack(
            side="right",
        )

        # --------------------------------------------------------
        # MAIN CONTENT
        # --------------------------------------------------------

        main_frame = ttk.Frame(
            self.root,
            padding=(15, 0),
        )

        main_frame.grid(
            row=3,
            column=0,
            sticky="nsew",
        )

        main_frame.columnconfigure(
            0,
            weight=1,
        )

        main_frame.rowconfigure(
            0,
            weight=1,
        )

        # --------------------------------------------------------
        # TABLE
        # --------------------------------------------------------

        table_frame = ttk.LabelFrame(
            main_frame,
            text=" Клиенты ",
            padding=5,
        )

        table_frame.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        table_frame.columnconfigure(
            0,
            weight=1,
        )

        table_frame.rowconfigure(
            0,
            weight=1,
        )

        columns = (
            "row",
            "date",
            "product",
            "name",
            "payment",
            "phone",
            "result",
        )

        self.table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )

        # Заголовки.
        self.table.heading(
            "row",
            text="Строка",
        )

        self.table.heading(
            "date",
            text="Дата",
        )

        self.table.heading(
            "product",
            text="Продукт",
        )

        self.table.heading(
            "name",
            text="Имя",
        )

        self.table.heading(
            "payment",
            text="Тип Оплаты",
        )

        self.table.heading(
            "phone",
            text="Номер",
        )

        # Только Result имеет фильтр.
        self.table.heading(
            "result",
            text="Результат ▾",
            command=self.open_result_filter,
        )

        # Ширина колонок.
        self.table.column(
            "row",
            width=65,
            minwidth=55,
            anchor="center",
        )

        self.table.column(
            "date",
            width=125,
            minwidth=90,
        )

        self.table.column(
            "product",
            width=300,
            minwidth=180,
        )

        self.table.column(
            "name",
            width=150,
            minwidth=100,
        )

        self.table.column(
            "payment",
            width=120,
            minwidth=90,
        )

        self.table.column(
            "phone",
            width=185,
            minwidth=140,
        )

        self.table.column(
            "result",
            width=145,
            minwidth=100,
        )

        # Текущая строка.
        self.table.tag_configure(
            "current",
            font=("Segoe UI", 9, "bold"),
        )

        scrollbar_y = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.table.yview,
        )

        scrollbar_x = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.table.xview,
        )

        self.table.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )

        self.table.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        scrollbar_y.grid(
            row=0,
            column=1,
            sticky="ns",
        )

        scrollbar_x.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        self.table.bind(
            "<<TreeviewSelect>>",
            self._on_table_select,
        )

        # --------------------------------------------------------
        # CURRENT CLIENT
        # --------------------------------------------------------

        current_frame = ttk.LabelFrame(
            main_frame,
            text=" Текущий клиент ",
            padding=20,
        )

        current_frame.grid(
            row=0,
            column=1,
            padx=(15, 0),
            sticky="ns",
        )

        self.current_row_label = ttk.Label(
            current_frame,
            text="Строка Excel: —",
            font=("Segoe UI", 11),
        )

        self.current_row_label.pack(
            pady=(0, 18),
        )

        self.current_phone = ttk.Label(
            current_frame,
            text="—",
            font=("Segoe UI", 19, "bold"),
        )

        self.current_phone.pack(
            pady=8,
        )

        self.current_name = ttk.Label(
            current_frame,
            text="—",
            font=("Segoe UI", 12),
        )

        self.current_name.pack(
            pady=5,
        )

        self.current_product = ttk.Label(
            current_frame,
            text="—",
            wraplength=300,
            justify="center",
        )

        self.current_product.pack(
            pady=10,
        )

        ttk.Separator(
            current_frame,
            orient="horizontal",
        ).pack(
            fill="x",
            pady=10,
        )

        self.current_result = ttk.Label(
            current_frame,
            text="Результат: —",
        )

        self.current_result.pack(
            pady=5,
        )

        self.attempt_label = ttk.Label(
            current_frame,
            text="Попытка: — / 3",
        )

        self.attempt_label.pack(
            pady=8,
        )

        self.call_state = ttk.Label(
            current_frame,
            text="READY",
            font=("Segoe UI", 12, "bold"),
        )

        self.call_state.pack(
            pady=8,
        )

        # --------------------------------------------------------
        # CONTROLS
        # --------------------------------------------------------

        control = ttk.Frame(
            self.root,
            padding=(15, 8),
        )

        control.grid(
            row=4,
            column=0,
            sticky="ew",
        )

        ttk.Label(
            control,
            text="Начать со строки:",
        ).pack(
            side="left",
            padx=(0, 8),
        )

        self.row_entry = ttk.Entry(
            control,
            width=8,
        )

        self.row_entry.pack(
            side="left",
        )

        ttk.Button(
            control,
            text="▶ Перейти",
            command=self.go_to_row,
        ).pack(
            side="left",
            padx=8,
        )

        ttk.Button(
            control,
            text="◀ Предыдущий",
            command=self.previous_row,
        ).pack(
            side="left",
            padx=5,
        )

        self.start_button = ttk.Button(
            control,
            text="▶ START",
            command=self.start_dialer,
        )

        self.start_button.pack(
            side="left",
            padx=(15, 5),
        )

        self.stop_button = ttk.Button(
            control,
            text="■ STOP",
            command=self.stop_dialer,
        )

        self.stop_button.pack(
            side="left",
            padx=5,
        )

        self.pause_button = ttk.Button(
            control,
            text="⏸ Пауза",
            command=self.toggle_pause,
        )

        self.pause_button.pack(
            side="left",
            padx=5,
        )

        ttk.Button(
            control,
            text="Следующий ▶",
            command=self.next_row,
        ).pack(
            side="left",
            padx=5,
        )

        self.wps_var = tk.BooleanVar(
            value=self.settings.wps_sync
        )

        ttk.Checkbutton(
            control,
            text="Синхронизация WPS",
            variable=self.wps_var,
            command=self.toggle_wps,
        ).pack(
            side="right",
        )

        # --------------------------------------------------------
        # LOG
        # --------------------------------------------------------

        log_frame = ttk.LabelFrame(
            self.root,
            text=" Лог ",
            padding=5,
        )

        log_frame.grid(
            row=5,
            column=0,
            padx=15,
            pady=(0, 15),
            sticky="nsew",
        )

        log_frame.columnconfigure(
            0,
            weight=1,
        )

        log_frame.rowconfigure(
            0,
            weight=1,
        )

        self.log = tk.Text(
            log_frame,
            height=8,
            state="disabled",
            font=("Consolas", 10),
        )

        self.log.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

    # ============================================================
    # FILE
    # ============================================================

    def select_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[
                (
                    "Excel files",
                    "*.xlsx *.xlsm",
                ),
                (
                    "All files",
                    "*.*",
                ),
            ],
        )

        if not path:
            return

        self.load_file(
            Path(path)
        )

    def load_file(
        self,
        path: Path,
    ) -> None:
        try:
            if self.controller is not None:
                self.controller.load_excel(path)

            reader = ExcelReader(path)
            rows = reader.load()

            if not rows:
                raise ValueError(
                    "В Excel не найдено ни одной строки с номером."
                )

            self.reader = reader
            self.all_rows = rows
            self.rows = list(rows)

            self.result_filter = ResultFilter(
                self.all_rows
            )

            self.file_label.config(
                text=str(path)
            )

            self.settings.last_excel_file = str(path)
            self.settings.last_excel_row = rows[0].excel_row
            self.settings.save()

            self.current_index = 0

            self.populate_table()
            self.select_current()
            self.update_filter_status()

            self.write_log(
                f"Файл загружен: {path.name}"
            )

            self.write_log(
                f"Найдено номеров: {len(rows)}"
            )

        except Exception as exc:
            messagebox.showerror(
                "Ошибка Excel",
                str(exc),
            )

    def _restore_last_file(self) -> None:
        path = Path(
            self.settings.last_excel_file
        )

        if not path.exists():
            return

        try:
            if self.controller is not None:
                self.controller.load_excel(path)

            reader = ExcelReader(path)
            rows = reader.load()

            if not rows:
                return

            self.reader = reader
            self.all_rows = rows
            self.rows = list(rows)

            self.result_filter = ResultFilter(
                self.all_rows
            )

            self.file_label.config(
                text=str(path)
            )

            last_row = self.settings.last_excel_row

            self.current_index = 0

            for index, row in enumerate(
                self.rows
            ):
                if row.excel_row >= last_row:
                    self.current_index = index
                    break

            self.populate_table()
            self.select_current()
            self.update_filter_status()

            self.write_log(
                f"Восстановлена последняя позиция: "
                f"строка {last_row}"
            )

            self.write_log(
                f"Файл загружен: {path.name}"
            )

            self.write_log(
                f"Найдено номеров: {len(rows)}"
            )

        except Exception as exc:
            self.write_log(
                f"Не удалось восстановить файл: {exc}"
            )

    # ============================================================
    # TABLE
    # ============================================================

    def populate_table(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)

        for index, row in enumerate(
            self.rows
        ):
            self.table.insert(
                "",
                "end",
                iid=f"row_{index}",
                values=(
                    row.excel_row,
                    row.date,
                    row.product,
                    row.name,
                    row.payment_type,
                    row.phone,
                    row.result,
                ),
            )

    def select_current(self) -> None:
        if not self.rows:
            self.current_index = -1

            self.current_row_label.config(
                text="Строка Excel: —"
            )

            self.current_phone.config(
                text="—"
            )

            self.current_name.config(
                text="—"
            )

            self.current_product.config(
                text="—"
            )

            self.current_result.config(
                text="Результат: —"
            )

            self.counter_label.config(
                text="0 / 0"
            )

            return

        if not (
            0 <= self.current_index < len(self.rows)
        ):
            self.current_index = 0

        row = self.rows[
            self.current_index
        ]

        item_id = (
            f"row_{self.current_index}"
        )

        # Убираем старое выделение.
        for item in self.table.get_children():
            self.table.item(
                item,
                tags=(),
            )

        # Выделяем текущую строку.
        self.table.item(
            item_id,
            tags=("current",),
        )

        self.table.selection_set(
            item_id
        )

        self.table.focus(
            item_id
        )

        # Прокручиваем таблицу к текущей строке.
        self.root.update_idletasks()

        self.table.see(
            item_id
        )

        # Правая панель.
        self.current_row_label.config(
            text=(
                f"Строка Excel: "
                f"{row.excel_row}"
            )
        )

        self.current_phone.config(
            text=row.phone
        )

        self.current_name.config(
            text=row.name or "—"
        )

        self.current_product.config(
            text=row.product or "—"
        )

        result = ResultFilter.normalize(
            row.result
        )

        self.current_result.config(
            text=f"Результат: {result}"
        )

        self.row_entry.delete(
            0,
            tk.END,
        )

        self.row_entry.insert(
            0,
            str(row.excel_row),
        )

        self.counter_label.config(
            text=(
                f"{self.current_index + 1} / "
                f"{len(self.rows)}"
            )
        )

        # Сохраняем прогресс.
        self.settings.last_excel_row = (
            row.excel_row
        )

        self.settings.save()

        # WPS.
        if self.wps_var.get():
            self.wps.select_row(
                row.excel_row
            )

    def _on_table_select(
        self,
        _event=None,
    ) -> None:
        selection = self.table.selection()

        if not selection:
            return

        item_id = selection[0]

        if not item_id.startswith(
            "row_"
        ):
            return

        try:
            index = int(
                item_id.replace(
                    "row_",
                    "",
                )
            )
        except ValueError:
            return

        if index == self.current_index:
            return

        self.current_index = index

        self.select_current()

    # ============================================================
    # RESULT FILTER
    # ============================================================

    def open_result_filter(self) -> None:
        if self.result_filter is None:
            messagebox.showinfo(
                "Фильтр",
                "Сначала выберите Excel-файл.",
            )
            return

        window = tk.Toplevel(
            self.root
        )

        window.title(
            "Фильтр — Результат"
        )

        window.resizable(
            False,
            False,
        )

        window.transient(
            self.root
        )

        window.grab_set()

        self.root.update_idletasks()

        # Открываем окно рядом с верхней частью интерфейса.
        x = (
            self.root.winfo_rootx()
            + 300
        )

        y = (
            self.root.winfo_rooty()
            + 160
        )

        window.geometry(
            f"340x450+{x}+{y}"
        )

        container = ttk.Frame(
            window,
            padding=12,
        )

        container.pack(
            fill="both",
            expand=True,
        )

        ttk.Label(
            container,
            text="Результат",
            font=("Segoe UI", 12, "bold"),
        ).pack(
            anchor="w",
            pady=(0, 5),
        )

        ttk.Label(
            container,
            text=(
                "Выберите значения, "
                "которые нужно показывать:"
            ),
        ).pack(
            anchor="w",
            pady=(0, 10),
        )

        values = (
            self.result_filter.get_values()
        )

        variables: dict[
            str,
            tk.BooleanVar,
        ] = {}

        current = (
            self.result_filter.selected
        )

        all_selected = (
            current is None
            or set(values) == current
        )

        select_all_var = tk.BooleanVar(
            value=all_selected
        )

        ttk.Checkbutton(
            container,
            text="Выбрать всё",
            variable=select_all_var,
            command=lambda: (
                self._toggle_all_result_values(
                    variables,
                    select_all_var.get(),
                )
            ),
        ).pack(
            anchor="w",
            pady=(0, 5),
        )

        ttk.Separator(
            container,
            orient="horizontal",
        ).pack(
            fill="x",
            pady=5,
        )

        # --------------------------------------------------------
        # VALUES
        # --------------------------------------------------------

        list_frame = ttk.Frame(
            container
        )

        list_frame.pack(
            fill="both",
            expand=True,
        )

        canvas = tk.Canvas(
            list_frame,
            width=290,
            height=245,
            highlightthickness=0,
        )

        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=canvas.yview,
        )

        values_frame = ttk.Frame(
            canvas
        )

        values_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(
                scrollregion=canvas.bbox(
                    "all"
                )
            ),
        )

        canvas.create_window(
            (0, 0),
            window=values_frame,
            anchor="nw",
        )

        canvas.configure(
            yscrollcommand=scrollbar.set
        )

        canvas.pack(
            side="left",
            fill="both",
            expand=True,
        )

        scrollbar.pack(
            side="right",
            fill="y",
        )

        for value in values:
            selected = (
                current is None
                or value in current
            )

            variable = tk.BooleanVar(
                value=selected
            )

            variables[value] = variable

            ttk.Checkbutton(
                values_frame,
                text=value,
                variable=variable,
            ).pack(
                anchor="w",
                pady=2,
            )

        # --------------------------------------------------------
        # BUTTONS
        # --------------------------------------------------------

        buttons = ttk.Frame(
            container
        )

        buttons.pack(
            fill="x",
            pady=(10, 0),
        )

        def reset() -> None:
            for variable in variables.values():
                variable.set(True)

            select_all_var.set(True)

        def apply() -> None:
            selected_values = {
                value
                for value, variable
                in variables.items()
                if variable.get()
            }

            if not selected_values:
                self.result_filter.set_selected(
                    set()
                )

            elif (
                selected_values
                == set(values)
            ):
                self.result_filter.clear()

            else:
                self.result_filter.set_selected(
                    selected_values
                )

            self.apply_filters(
                preserve_current=True
            )

            window.destroy()

        ttk.Button(
            buttons,
            text="Сбросить",
            command=reset,
        ).pack(
            side="left",
        )

        ttk.Button(
            buttons,
            text="Применить",
            command=apply,
        ).pack(
            side="right",
        )

    def _toggle_all_result_values(
        self,
        variables: dict[
            str,
            tk.BooleanVar,
        ],
        value: bool,
    ) -> None:
        for variable in variables.values():
            variable.set(value)

    # ============================================================
    # SEARCH
    # ============================================================

    def _on_search(
        self,
        _event=None,
    ) -> None:
        self.search_text = (
            self.search_var
            .get()
            .strip()
            .lower()
        )

        self.apply_filters(
            preserve_current=True
        )

    def clear_search(self) -> None:
        self.search_var.set("")
        self.search_text = ""

        self.apply_filters(
            preserve_current=True
        )

    def _matches_search(
        self,
        row: ExcelRow,
    ) -> bool:
        if not self.search_text:
            return True

        values = (
            row.date,
            row.product,
            row.name,
            row.payment_type,
            row.phone,
            row.result,
        )

        return any(
            self.search_text
            in str(value).lower()
            for value in values
        )

    # ============================================================
    # APPLY FILTERS
    # ============================================================

    def apply_filters(
        self,
        preserve_current: bool = True,
    ) -> None:
        if self.result_filter is None:
            return

        old_excel_row = None

        if (
            preserve_current
            and 0 <= self.current_index
            < len(self.rows)
        ):
            old_excel_row = (
                self.rows[
                    self.current_index
                ].excel_row
            )

        # Result filter.
        rows = self.result_filter.apply(
            self.all_rows
        )

        # Search.
        if self.search_text:
            rows = [
                row
                for row in rows
                if self._matches_search(row)
            ]

        self.rows = rows

        # Ничего не найдено.
        if not self.rows:
            self.current_index = -1

            self.populate_table()
            self.select_current()
            self.update_filter_status()

            self.write_log(
                "Фильтр: подходящих клиентов нет."
            )

            return

        # По умолчанию первая строка.
        self.current_index = 0

        # Если возможно — сохраняем текущую строку.
        if old_excel_row is not None:
            for index, row in enumerate(
                self.rows
            ):
                if (
                    row.excel_row
                    == old_excel_row
                ):
                    self.current_index = index
                    break

        self.populate_table()
        self.select_current()
        self.update_filter_status()

        self.write_log(
            f"Фильтр применён: "
            f"{len(self.rows)} из "
            f"{len(self.all_rows)}"
        )

    def clear_all_filters(self) -> None:
        if self.result_filter is not None:
            self.result_filter.clear()

        self.search_var.set("")
        self.search_text = ""

        self.rows = list(
            self.all_rows
        )

        self.current_index = 0

        self.populate_table()
        self.select_current()
        self.update_filter_status()

        self.write_log(
            "Фильтр результата и поиск сброшены."
        )

    def update_filter_status(self) -> None:
        if not self.all_rows:
            self.filter_status.config(
                text="Показано: 0"
            )
            return

        active = []

        if (
            self.result_filter is not None
            and self.result_filter.is_active
        ):
            active.append(
                "Результат"
            )

        if self.search_text:
            active.append(
                "Поиск"
            )

        if active:
            self.filter_status.config(
                text=(
                    f"{len(self.rows)} / "
                    f"{len(self.all_rows)} | "
                    f"{', '.join(active)}"
                )
            )

            self.result_filter_button.config(
                text="Результат ●"
            )

        else:
            self.filter_status.config(
                text=(
                    f"{len(self.rows)} / "
                    f"{len(self.all_rows)}"
                )
            )

            self.result_filter_button.config(
                text="Результат ▾"
            )

    # ============================================================
    # NAVIGATION
    # ============================================================

    def go_to_row(self) -> None:
        try:
            excel_row = int(
                self.row_entry.get().strip()
            )
        except ValueError:
            messagebox.showwarning(
                "Строка",
                "Введите номер строки.",
            )
            return

        for index, row in enumerate(
            self.rows
        ):
            if row.excel_row == excel_row:
                self.current_index = index

                self.select_current()

                self.write_log(
                    f"Переход → строка "
                    f"{excel_row}"
                )

                return

        messagebox.showwarning(
            "Строка",
            (
                f"Строка {excel_row} "
                "не найдена среди "
                "текущих результатов."
            ),
        )

    def previous_row(self) -> None:
        if not self.rows:
            return

        if self.current_index <= 0:
            self.write_log(
                "Уже первая доступная строка."
            )
            return

        self.current_index -= 1

        self.select_current()

        self.write_log(
            f"Назад → строка "
            f"{self.rows[self.current_index].excel_row}"
        )

    def next_row(self) -> None:
        if not self.rows:
            return

        if (
            self.current_index
            >= len(self.rows) - 1
        ):
            self.write_log(
                "Достигнут конец текущего списка."
            )
            return

        self.current_index += 1

        self.select_current()

        self.write_log(
            f"Вперёд → строка "
            f"{self.rows[self.current_index].excel_row}"
        )

    # ============================================================
    # START / PAUSE / STOP
    # ============================================================

    def _sync_controller_filter(self) -> None:
        if self.controller is None:
            return

        if self.result_filter is None:
            self.controller.clear_result_filter()
            return

        if self.result_filter.is_active:
            selected = self.result_filter.selected
            self.controller.set_result_filter(
                None if selected is None else set(selected)
            )
        else:
            self.controller.clear_result_filter()

    def _get_start_row(self) -> int | None:
        value = self.row_entry.get().strip()

        if not value:
            if (
                0 <= self.current_index < len(self.rows)
            ):
                return self.rows[self.current_index].excel_row
            return None

        try:
            return int(value)
        except ValueError:
            raise ValueError(
                "Начальная строка должна быть целым числом."
            )

    def start_dialer(self) -> None:
        if self.controller is None:
            messagebox.showerror(
                "ZIK",
                "Controller не подключён.",
            )
            return

        if not self.all_rows:
            messagebox.showwarning(
                "ZIK",
                "Сначала выберите Excel-файл.",
            )
            return

        try:
            start_row = self._get_start_row()
            self._sync_controller_filter()

            count = self.controller.prepare_queue(
                start_row=start_row,
            )

            if count <= 0:
                messagebox.showwarning(
                    "ZIK",
                    "После фильтрации очередь пуста.",
                )
                return

            self.controller.start(
                start_row=start_row,
            )

            self.paused = False

            self.start_button.config(
                state="disabled",
            )

            self.stop_button.config(
                state="normal",
            )

            self.pause_button.config(
                text="⏸ Пауза",
                state="normal",
            )

            self.status_label.config(
                text="● RUNNING",
            )

            self.write_log(
                f"▶ START: очередь запущена ({count} номеров)."
            )

        except Exception as exc:
            messagebox.showerror(
                "Ошибка запуска",
                str(exc),
            )
            self.write_log(
                f"START ERROR: {exc}"
            )

    def stop_dialer(self) -> None:
        if self.controller is None:
            return

        try:
            self.controller.stop()
        except Exception as exc:
            self.write_log(
                f"STOP ERROR: {exc}"
            )
            return

        self.paused = False

        self.start_button.config(
            state="normal",
        )

        self.stop_button.config(
            state="normal",
        )

        self.pause_button.config(
            text="⏸ Пауза",
            state="normal",
        )

        self.status_label.config(
            text="● READY",
        )

        self.call_state.config(
            text="READY",
        )

        self.write_log(
            "■ STOP: автодозвон остановлен."
        )

    def toggle_pause(self) -> None:
        if self.controller is None:
            self.write_log(
                "PAUSE: Controller не подключён."
            )
            return

        if not self.controller.is_running:
            self.write_log(
                "PAUSE: автодозвон сейчас не запущен."
            )
            return

        if self.controller.is_paused:
            self.controller.resume()

            self.paused = False

            self.pause_button.config(
                text="⏸ Пауза"
            )

            self.status_label.config(
                text="● RUNNING"
            )

            self.write_log(
                "▶ RESUME: автодозвон продолжен."
            )

        else:
            self.controller.pause()

            self.paused = True

            self.pause_button.config(
                text="▶ Продолжить"
            )

            self.status_label.config(
                text="● PAUSED"
            )

            self.write_log(
                "⏸ PAUSE: текущий звонок НЕ завершается."
            )

            self.write_log(
                "Следующий номер будет набран после RESUME."
            )

    def _on_controller_status(
        self,
        status: dict,
    ) -> None:
        """
        Callback Controller может приходить из фонового потока.
        Tkinter нельзя обновлять из worker thread, поэтому
        всегда передаём обновление через root.after().
        """
        try:
            self.root.after(
                0,
                self._apply_controller_status,
                status,
            )
        except tk.TclError:
            pass

    def _apply_controller_status(
        self,
        status: dict,
    ) -> None:
        state = status.get("state", "IDLE")
        phone = status.get("phone") or ""
        excel_row = status.get("excel_row")
        attempt = status.get("attempt", 0)
        result = status.get("result") or ""
        running = bool(status.get("running"))
        paused = bool(status.get("paused"))

        self.call_state.config(
            text=(
                "PAUSED"
                if paused
                else state
            )
        )

        if paused:
            self.status_label.config(
                text="● PAUSED"
            )
        elif running:
            self.status_label.config(
                text="● RUNNING"
            )
        else:
            self.status_label.config(
                text="● READY"
            )

        self.attempt_label.config(
            text=f"Попытка: {attempt} / 3"
        )

        if phone:
            self.current_phone.config(
                text=phone
            )

        if excel_row is not None:
            self.current_row_label.config(
                text=f"Строка Excel: {excel_row}"
            )

        if result:
            self.current_result.config(
                text=f"Результат: {result}"
            )

        queue_size = status.get("queue_size", 0)

        if running:
            self.start_button.config(
                state="disabled"
            )
            self.stop_button.config(
                state="normal"
            )
        else:
            self.start_button.config(
                state="normal"
            )

        self.write_log(
            f"STATE: {state}"
            + (
                f" | {phone}"
                if phone
                else ""
            )
            + (
                f" | row={excel_row}"
                if excel_row is not None
                else ""
            )
            + (
                f" | attempt={attempt}/3"
                if attempt
                else ""
            )
            + (
                f" | result={result}"
                if result
                else ""
            )
            + f" | queue={queue_size}"
        )

    # ============================================================
    # ============================================================
    # WPS
    # ============================================================

    def toggle_wps(self) -> None:
        self.wps.enabled = (
            self.wps_var.get()
        )

        self.settings.wps_sync = (
            self.wps_var.get()
        )

        self.settings.save()

        self.write_log(
            "WPS sync: "
            + (
                "ON"
                if self.wps.enabled
                else "OFF"
            )
        )

    # ============================================================
    # LOG
    # ============================================================

    def _on_controller_log(
        self,
        message: str,
    ) -> None:
        try:
            self.root.after(
                0,
                self.write_log,
                message,
            )
        except tk.TclError:
            pass

    def write_log(
        self,
        message: str,
    ) -> None:
        self.log.config(
            state="normal"
        )

        self.log.insert(
            tk.END,
            message + "\n",
        )

        self.log.see(
            tk.END
        )

        self.log.config(
            state="disabled"
        )

    # ============================================================
    # RUN
    # ============================================================

    def run(self) -> None:
        self.root.mainloop()