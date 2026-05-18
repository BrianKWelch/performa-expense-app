def calculate_per_diem_total(
    days: int,
    *,
    first_last_rate: float = 75.0,
    middle_rate: float = 100.0,
) -> float:
    if days <= 0:
        return 0.0
    if days == 1:
        return first_last_rate
    if days == 2:
        return first_last_rate * 2
    return first_last_rate * 2 + (days - 2) * middle_rate


def format_per_diem_summary(
    days: int,
    *,
    first_last_rate: float = 75.0,
    middle_rate: float = 100.0,
) -> str:
    total = calculate_per_diem_total(
        days, first_last_rate=first_last_rate, middle_rate=middle_rate
    )
    if days <= 0:
        return f"Per diem total ${total:.2f}"
    if days == 1:
        return f"Per diem: 1 day @ ${first_last_rate:.0f} = ${total:.2f}"
    if days == 2:
        return (
            f"Per diem: 2 days @ ${first_last_rate:.0f} each (first and last) = ${total:.2f}"
        )
    middle_days = days - 2
    return (
        f"Per diem: ${first_last_rate:.0f} first day, ${middle_rate:.0f} × {middle_days} "
        f"middle day(s), ${first_last_rate:.0f} last day = ${total:.2f}"
    )
