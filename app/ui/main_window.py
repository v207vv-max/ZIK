from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.excel.filter import ColumnFilter, ResultFilter
from app.excel.reader import ExcelReader, ExcelRow
from app.settings import Settings


class MainWindow:
    def __init__(self, controller=None) -> None:
        self.controller = controller

        self.root = tk.Tk()

        self.root.title("ZIK — Smart Call Center")
        self.root.geometry("1450x850")
        self.root.minsize(1150, 700)

        self.settings = Settings()
        self.reader: ExcelReader | None = None
        self.result_filter: ResultFilter | None = None
        self.status_filter: set[str] | None = None
        self.column_filters: dict[str, set[str] | None] = {}

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

        # Keep the visible ZIK table synchronized with the active
        # Dialer item. Controller callbacks arrive from worker threads,
        # so the UI also polls controller state from the Tk main thread.
        self._controller_poll_job = self.root.after(
            100,
            self._poll_controller_status,
        )

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
            text="Фильтры",
            command=self.open_all_filters,
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
            "status",
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

        # Колонка F из Excel.
        self.table.heading(
            "status",
            text="Статус ▾",
            command=lambda: self.open_column_filter("status", "Статус"),
        )

        # Колонка G из Excel.
        self.table.heading(
            "result",
            text="Результат ▾",
            command=lambda: self.open_column_filter("result", "Результат"),
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
            "status",
            width=135,
            minwidth=100,
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
            text="Попытка: — / 2",
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
            self.status_filter = None
            self.column_filters.clear()

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
            self.status_filter = None
            self.column_filters.clear()

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
                    row.status,
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
    # COLUMN FILTERS
    # ============================================================

    # Только полезные фильтры по значениям.
    # «Строка», «Дата» и «Номер» не являются checkbox-фильтрами:
    # строка используется для навигации, дата и номер обычно уникальны.
    _FILTER_COLUMNS = (
        ("status", "Статус"),
        ("result", "Результат"),
    )

    @staticmethod
    def _column_label(column: str) -> str:
        return dict(MainWindow._FILTER_COLUMNS).get(column, column)

    @staticmethod
    def _column_value(column: str, row: ExcelRow) -> str:
        if column == "row":
            return str(row.excel_row)
        if column == "date":
            return ColumnFilter.normalize(row.date)
        if column == "product":
            return ColumnFilter.normalize(row.product)
        if column == "name":
            return ColumnFilter.normalize(row.name)
        if column == "payment":
            return ColumnFilter.normalize(row.payment_type)
        if column == "phone":
            return ColumnFilter.normalize(row.phone)
        if column == "status":
            return ColumnFilter.normalize(row.status)
        if column == "result":
            return ResultFilter.normalize(row.result)
        return ""

    def _get_column_values(self, column: str) -> list[str]:
        values = {
            ColumnFilter.display_value(
                self._column_value(column, row)
            )
            for row in self.all_rows
        }

        return sorted(
            values,
            key=lambda value: (
                value == ColumnFilter.EMPTY_LABEL,
                value.casefold(),
            ),
        )

    def _get_active_filter(self, column: str) -> set[str] | None:
        selected = self.column_filters.get(column)

        if column == "result" and self.result_filter is not None:
            selected = self.result_filter.selected

        return None if selected is None else set(selected)

    def open_all_filters(self) -> None:
        # Основная кнопка «Фильтры» открывает фильтр Статус.
        self.open_column_filter("status", "Статус")

    def open_column_filter(
        self,
        column: str,
        label: str | None = None,
        parent=None,
    ) -> None:
        if not self.all_rows:
            messagebox.showinfo(
                "Фильтры",
                "Сначала выберите Excel-файл.",
            )
            return

        window = tk.Toplevel(self.root)
        window.title("Фильтры")
        window.geometry("470x650")
        window.minsize(440, 560)
        window.resizable(True, True)
        window.transient(parent or self.root)
        window.grab_set()

        container = ttk.Frame(window, padding=12)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Фильтр по колонке",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        labels = [name for _, name in self._FILTER_COLUMNS]
        column_var = tk.StringVar(
            value=label or self._column_label(column)
        )

        column_combo = ttk.Combobox(
            container,
            textvariable=column_var,
            values=labels,
            state="readonly",
        )
        column_combo.pack(fill="x", pady=(0, 8))

        search_var = tk.StringVar()

        search_frame = ttk.Frame(container)
        search_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(
            search_frame,
            text="Поиск:",
        ).pack(side="left", padx=(0, 8))

        search_entry = ttk.Entry(
            search_frame,
            textvariable=search_var,
        )
        search_entry.pack(side="left", fill="x", expand=True)

        list_frame = ttk.Frame(container)
        list_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            list_frame,
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=canvas.yview,
        )
        values_frame = ttk.Frame(canvas)

        canvas_window = canvas.create_window(
            (0, 0),
            window=values_frame,
            anchor="nw",
        )

        values_frame.bind(
            "<Configure>",
            lambda _event: canvas.configure(
                scrollregion=canvas.bbox("all")
            ),
        )

        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(
                canvas_window,
                width=event.width,
            ),
        )

        canvas.configure(
            yscrollcommand=scrollbar.set,
        )

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        current_column = {"key": column}
        variables: dict[str, tk.BooleanVar] = {}
        rebuilding = {"value": False}

        select_all_var = tk.BooleanVar(value=True)

        def update_select_all() -> None:
            if not variables:
                select_all_var.set(False)
                return
            select_all_var.set(
                all(variable.get() for variable in variables.values())
            )

        def rebuild_values() -> None:
            rebuilding["value"] = True

            for child in values_frame.winfo_children():
                child.destroy()

            variables.clear()

            active_column = current_column["key"]
            selected = self._get_active_filter(active_column)
            query = search_var.get().strip().casefold()

            for display_value in self._get_column_values(active_column):
                if query and query not in display_value.casefold():
                    continue

                stored_value = ColumnFilter.to_stored(display_value)
                checked = (
                    selected is None
                    or stored_value in selected
                )

                variable = tk.BooleanVar(value=checked)
                variables[stored_value] = variable

                ttk.Checkbutton(
                    values_frame,
                    text=display_value,
                    variable=variable,
                    command=update_select_all,
                ).pack(
                    anchor="w",
                    fill="x",
                    pady=1,
                )

            update_select_all()
            rebuilding["value"] = False

        def change_column(_event=None) -> None:
            selected_label = column_var.get()

            for key, name in self._FILTER_COLUMNS:
                if name == selected_label:
                    current_column["key"] = key
                    break

            search_var.set("")
            rebuild_values()

        column_combo.bind(
            "<<ComboboxSelected>>",
            change_column,
        )

        def toggle_all() -> None:
            value = select_all_var.get()
            for variable in variables.values():
                variable.set(value)

        ttk.Checkbutton(
            container,
            text="Выбрать всё",
            variable=select_all_var,
            command=toggle_all,
        ).pack(anchor="w", pady=(0, 6))

        search_var.trace_add(
            "write",
            lambda *_args: rebuild_values()
            if not rebuilding["value"]
            else None,
        )

        quick_buttons = ttk.Frame(container)
        quick_buttons.pack(fill="x", pady=(6, 8))

        def select_visible() -> None:
            select_all_var.set(True)
            for variable in variables.values():
                variable.set(True)

        def deselect_visible() -> None:
            select_all_var.set(False)
            for variable in variables.values():
                variable.set(False)

        ttk.Button(
            quick_buttons,
            text="Все",
            command=select_visible,
        ).pack(side="left")

        ttk.Button(
            quick_buttons,
            text="Ничего",
            command=deselect_visible,
        ).pack(side="left", padx=5)

        ttk.Button(
            quick_buttons,
            text="Сбросить все",
            command=self.clear_all_filters,
        ).pack(side="right")

        buttons = ttk.Frame(container)
        buttons.pack(fill="x", pady=(4, 0))

        def apply() -> None:
            active_column = current_column["key"]

            all_values = {
                ColumnFilter.to_stored(value)
                for value in self._get_column_values(active_column)
            }

            selected_values = {
                value
                for value, variable in variables.items()
                if variable.get()
            }

            # Если поле поиска использовалось, его скрытые значения
            # должны остаться выбранными. Иначе поиск значений мог бы
            # случайно удалить уже установленный фильтр.
            query = search_var.get().strip().casefold()

            if query:
                previous = self._get_active_filter(active_column)

                if previous is None:
                    hidden = {
                        value
                        for value in all_values
                        if query not in ColumnFilter.display_value(value).casefold()
                    }
                    selected_values |= hidden
                else:
                    selected_values |= (
                        previous - set(variables)
                    )

            if active_column == "result":
                selected_values = {
                    ResultFilter.normalize(value)
                    for value in selected_values
                }
                all_values = {
                    ResultFilter.normalize(value)
                    for value in all_values
                }

            if not selected_values or selected_values == all_values:
                self.column_filters.pop(active_column, None)

                if (
                    active_column == "result"
                    and self.result_filter is not None
                ):
                    self.result_filter.clear()
            else:
                self.column_filters[active_column] = selected_values

                if (
                    active_column == "result"
                    and self.result_filter is not None
                ):
                    self.result_filter.set_selected(selected_values)

            self.apply_filters(preserve_current=True)
            window.destroy()

        ttk.Button(
            buttons,
            text="Отмена",
            command=window.destroy,
        ).pack(side="left")

        ttk.Button(
            buttons,
            text="Применить",
            command=apply,
        ).pack(side="right")

        rebuild_values()
        search_entry.focus_set()

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
            row.status,
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
        old_excel_row = None

        if (
            preserve_current
            and 0 <= self.current_index < len(self.rows)
        ):
            old_excel_row = self.rows[self.current_index].excel_row

        rows = list(self.all_rows)

        for column, selected in self.column_filters.items():
            if selected is not None:
                rows = [
                    row
                    for row in rows
                    if self._column_value(column, row) in selected
                ]

        if self.search_text:
            rows = [
                row for row in rows
                if self._matches_search(row)
            ]

        self.rows = rows

        if not self.rows:
            self.current_index = -1
            self.populate_table()
            self.select_current()
            self.update_filter_status()
            self.write_log("Фильтр: подходящих клиентов нет.")
            return

        self.current_index = 0

        if old_excel_row is not None:
            for index, row in enumerate(self.rows):
                if row.excel_row == old_excel_row:
                    self.current_index = index
                    break

        self.populate_table()
        self.select_current()
        self.update_filter_status()

    def clear_all_filters(self) -> None:
        if self.result_filter is not None:
            self.result_filter.clear()

        self.status_filter = None
        self.column_filters.clear()

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
            "Все фильтры и поиск сброшены."
        )

    def update_filter_status(self) -> None:
        if not self.all_rows:
            self.filter_status.config(text="Показано: 0")
            return

        active = [
            self._column_label(column)
            for column, selected in self.column_filters.items()
            if selected is not None
        ]

        if self.search_text:
            active.append("Поиск")

        if active:
            self.filter_status.config(
                text=(
                    f"{len(self.rows)} / {len(self.all_rows)} | "
                    f"{', '.join(active)}"
                )
            )
            self.result_filter_button.config(text="Фильтры ●")
        else:
            self.filter_status.config(
                text=f"{len(self.rows)} / {len(self.all_rows)}"
            )
            self.result_filter_button.config(text="Фильтры")

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

        selected = self.column_filters.get("result")

        if selected is not None:
            self.controller.set_result_filter(set(selected))
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

    def _poll_controller_status(self) -> None:
        """
        Refresh the active call information from the Tk main thread.

        This is intentionally lightweight and does not write to the log.
        It prevents the visible table selection from waiting for a later
        Tkinter event when Dialer changes the current queue item.
        """
        try:
            if self.controller is not None:
                running = self.controller.is_running
                paused = self.controller.is_paused

                if running or paused:
                    current = self.controller.get_current()
                    excel_row = current.get("excel_row")
                    phone = current.get("phone") or ""
                    state = current.get("state") or "IDLE"
                    attempt = current.get("attempt", 0)
                    result = current.get("result") or ""

                    if excel_row is not None:
                        self._sync_ui_to_excel_row(
                            int(excel_row)
                        )

                        self.current_row_label.config(
                            text=f"Строка Excel: {excel_row}"
                        )

                    if phone:
                        self.current_phone.config(
                            text=phone
                        )

                    self.attempt_label.config(
                        text=f"Попытка: {attempt} / 2"
                    )

                    if result:
                        self.current_result.config(
                            text=f"Результат: {result}"
                        )

                    self.call_state.config(
                        text="PAUSED" if paused else state
                    )

        except (tk.TclError, RuntimeError):
            return
        except Exception:
            return

        try:
            self._controller_poll_job = self.root.after(
                100,
                self._poll_controller_status,
            )
        except tk.TclError:
            pass

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

    def _sync_ui_to_excel_row(
        self,
        excel_row: int,
    ) -> None:
        """
        Synchronize the visible UI selection with the Excel row
        reported by Controller.
        """
        for index, row in enumerate(self.rows):
            if row.excel_row == excel_row:
                if index != self.current_index:
                    self.current_index = index
                    self.select_current()
                return

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
            text=f"Попытка: {attempt} / 2"
        )

        if phone:
            self.current_phone.config(
                text=phone
            )

        if excel_row is not None:
            self._sync_ui_to_excel_row(
                int(excel_row)
            )

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
                f" | attempt={attempt}/2"
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
