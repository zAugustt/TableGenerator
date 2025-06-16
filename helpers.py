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

def change_header_order(headers: list[str], values: list[float]):
    """
    Re-orders the lists to adjust for cases in which we want totals at the bottom or top.
    """