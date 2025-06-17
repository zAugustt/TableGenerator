def format_headers(headers: list[str]) -> list[str]:
    """
    Formats the headers to have capitalized first letters on each word.
    :param headers: List of the respective headers
    """
    headers = [header.title() for header in headers if isinstance(header, str)]
    return headers


def format_values(values: list[float]) -> list[float]:
    """
    Formats the values to be rounded based on the common rounding rule and appearing as whole percentages.
    :param values: List of the respective values
    """
    values = [round(value * 100) if isinstance(value, (int, float)) else value for value in values]
    return values


def add_percentages_to_values(value: str) -> str:
    """
    TODO: Docstring
    """
    percentage_dict = ["--", "*"]
    if value in percentage_dict:
        return str(value)
    elif value is not None:
        return str(value) + "%"
    else:
        return ""

# Copilot assisted with the list comprehension here
def move_totals_to_bottom(headers: list[str], values: list[str]) -> tuple[list[str], list[str]]:
    """
    Moves headers containing "total" and their corresponding values to the bottom of the lists.
    :param headers: List of headers.
    :param values: List of values corresponding to the headers.
    :return: Tuple of reordered headers and values.
    """
    # Separate "total" headers and their values
    total_items = [(header, value) for header, value in zip(headers, values) if "total" in header.lower()]
    non_total_items = [(header, value) for header, value in zip(headers, values) if "total" not in header.lower()]

    # Combine non-total items with total items at the end
    reordered_headers, reordered_values = zip(*(non_total_items + total_items)) if non_total_items or total_items else ([], [])
    return list(reordered_headers), list(reordered_values)
