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
    Adds a percentage symbol to the given value unless it is in the predefined list of exceptions.

    :param value: The value to be formatted as a percentage.
    :return: The formatted value as a string with a percentage symbol, or the original value if it is in the exception list.
    """
    percentage_dict = ["--", "*"]
    if str(value).isnumeric():
        return str(value) + "%"
    elif value is not None or value in percentage_dict:
        return str(value)
    else:
        return ""


# Copilot assisted with the list comprehension here
# def move_totals(headers: list[str], values: list[str],direction: str) -> tuple[list[str], list[str]]:
#     """
#    Moves headers containing "total" and their corresponding values to the bottom of the lists.
#    :param headers: List of headers.
#    :param values: List of values corresponding to the headers.
#    :param direction: Defines which direction to move the totals (top/bottom)
#    :return: Tuple of reordered headers and values.
#    """
#
#     # Separate "total" headers and their values
#     ds = [(header, value) for header, value in zip(headers, values) if "**d/s" in header.lower()]
#     total_items = [(header, value) for header, value in zip(headers, values) if "total" in header.lower()]
#     non_total_items = [(header, value) for header, value in zip(headers, values) if
#                        "total" not in header.lower() and "**d/s" not in header.lower()]
#
#     if direction == "Bottom":
#         reordered_headers, reordered_values = zip(
#             *(ds + non_total_items + total_items)) if non_total_items or total_items else ([], [])
#     else:
#         reordered_headers, reordered_values = zip(
#             *(ds + total_items + non_total_items)) if total_items or non_total_items else ([], [])
#
#     return list(reordered_headers), list(reordered_values)

def move_totals(headers: list[str], values: list[list[str]], direction: str) -> tuple[list[str], list[list[str]]]:
    """
    Moves headers containing "total" and their corresponding values to the bottom of the lists.
    Subsets remain unchanged.
    :param headers: List of headers.
    :param values: List of value arrays corresponding to the headers.
    :param direction: Defines which direction to move the totals (top/bottom).
    :return: Tuple of reordered headers, values, and subsets.
    """

    # Separate "total" headers and their values
    ds_indices = [i for i, header in enumerate(headers) if "**d/s" in header.lower()]
    total_indices = [i for i, header in enumerate(headers) if "total" in header.lower()]
    non_total_indices = [i for i, header in enumerate(headers) if
                         "total" not in header.lower() and "**d/s" not in header.lower()]

    if direction == "Bottom":
        reordered_indices = ds_indices + non_total_indices + total_indices
    else:
        reordered_indices = ds_indices + total_indices + non_total_indices

    # Reorder headers and values based on the indices
    reordered_headers = [headers[i] for i in reordered_indices]
    reordered_values = [[value[i] for i in reordered_indices] for value in values]

    return reordered_headers, reordered_values
