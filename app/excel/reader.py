from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import openpyxl


@dataclass
class ExcelRow:
    excel_row: int
    date: object
    product: object
    name: object
    payment_type: object
    phone: object
    status: object = ""
    result: object = ""


class ExcelReader:
    REQUIRED_COLUMNS = {
        "Дата": "date",
        "Продукт": "product",
        "Имя": "name",
        "Тип Оплаты": "payment_type",
        "Номер": "phone",
    }

    STATUS_COLUMN_INDEX = 6
    RESULT_COLUMN = "Результат"

    def __init__(self, path: Path) -> None:
        self.path = Path(path)

    @staticmethod
    def _normalize_header(value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def _find_result_column(
        self,
        worksheet,
        headers: dict[str, int],
    ) -> int | None:
        result_column = headers.get(self.RESULT_COLUMN)

        if result_column is not None:
            return result_column

        required_columns = [
            headers[column]
            for column in self.REQUIRED_COLUMNS
        ]

        if not required_columns:
            return None

        next_column = max(required_columns) + 1

        if next_column > worksheet.max_column:
            return None

        for row_number in range(2, worksheet.max_row + 1):
            value = worksheet.cell(
                row=row_number,
                column=next_column,
            ).value

            if value is not None and str(value).strip():
                return next_column

        return None

    def load(self) -> list[ExcelRow]:
        if not self.path.exists():
            raise FileNotFoundError(
                f"Excel-файл не найден: {self.path}"
            )

        workbook = openpyxl.load_workbook(
            self.path,
            data_only=False,
        )

        try:
            worksheet = workbook.active

            headers: dict[str, int] = {}

            for index, cell in enumerate(
                worksheet[1],
                start=1,
            ):
                header = self._normalize_header(cell.value)

                if header:
                    headers[header] = index

            missing = [
                column
                for column in self.REQUIRED_COLUMNS
                if column not in headers
            ]

            if missing:
                raise ValueError(
                    "В Excel отсутствуют "
                    f"обязательные колонки: {missing}"
                )

            result_column = self._find_result_column(
                worksheet,
                headers,
            )

            status_column = (
                max(headers[column] for column in self.REQUIRED_COLUMNS)
                + 1
            )

            rows: list[ExcelRow] = []

            for row_number in range(
                2,
                worksheet.max_row + 1,
            ):
                status = worksheet.cell(
                    row=row_number,
                    column=status_column,
                ).value if status_column <= worksheet.max_column else ""

                result = ""
                if result_column is not None:
                    result = worksheet.cell(
                        row=row_number,
                        column=result_column,
                    ).value

                rows.append(
                    ExcelRow(
                        excel_row=row_number,
                        date=worksheet.cell(
                            row=row_number,
                            column=headers["Дата"],
                        ).value,
                        product=worksheet.cell(
                            row=row_number,
                            column=headers["Продукт"],
                        ).value,
                        name=worksheet.cell(
                            row=row_number,
                            column=headers["Имя"],
                        ).value,
                        payment_type=worksheet.cell(
                            row=row_number,
                            column=headers["Тип Оплаты"],
                        ).value,
                        phone=worksheet.cell(
                            row=row_number,
                            column=headers["Номер"],
                        ).value,
                        status=status,
                        result=result,
                    )
                )

            return rows

        finally:
            workbook.close()
