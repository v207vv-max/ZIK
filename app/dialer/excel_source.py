from __future__ import annotations

import logging

from app.excel.phone import normalize_phone
from app.excel.reader import ExcelRow


logger = logging.getLogger(
    "ZIK.ExcelSource"
)


def build_call_items(
    rows: list[ExcelRow],
) -> list[tuple[str, int]]:
    """
    Превращает строки Excel в очередь звонков.

    Возвращает:
        [
            ("950202495", 62),
            ("939405884", 64),
            ...
        ]

    где второе значение — настоящая строка Excel.
    """

    items: list[tuple[str, int]] = []

    for row in rows:
        try:
            phone = normalize_phone(
                row.phone
            )

        except ValueError as exc:
            logger.warning(
                "EXCEL: пропущен некорректный номер | "
                "row=%s phone=%r error=%s",
                row.excel_row,
                row.phone,
                exc,
            )

            continue

        if not phone:
            logger.warning(
                "EXCEL: пропущена пустая строка | "
                "row=%s",
                row.excel_row,
            )

            continue

        items.append(
            (
                phone,
                row.excel_row,
            )
        )

    logger.info(
        "EXCEL: prepared call queue | "
        "items=%s",
        len(items),
    )

    return items
