from __future__ import annotations

import re


def normalize_phone(value: object) -> str:
    """
    Приводит номер из Excel к формату MicroSIP.

    +998950202495
    998950202495
    +998 95 020 24 95

    ->

    950202495
    """

    if value is None:
        return ""

    phone = str(value).strip()

    if not phone:
        return ""

    # Оставляем только цифры.
    digits = re.sub(
        r"\D",
        "",
        phone,
    )

    # Узбекистан:
    # 998 + 9 цифр = 12 цифр.
    if len(digits) == 12 and digits.startswith("998"):
        digits = digits[3:]

    # Уже девятизначный номер.
    if len(digits) == 9:
        return digits

    raise ValueError(
        f"Некорректный номер телефона: {value!r}"
    )