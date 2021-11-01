"""
Utility functions for use with XlClass mudule.
"""


from openpyxl.utils import get_column_letter


def generate_columns_dictionary(key_list: list) -> dict:
    """Uses the passed ordered list (key_list) of values to generate a
    dictionary of corresponding column letters.
    {list0: "A", list1: "B", list2: "C"}

    Args:
        key_list (list): Ordered list of values to be used as the keys
        in the generated dictionary.

    Returns:
        dict: Dictionary of list items and their corresponding column
        letters. {list0: "A", list1: "B", list2: "C"}
    """
    return {key_value: get_column_letter(
        column_number) for column_number, key_value in enumerate(key_list, 1)}


def generate_source_target_columns_dictionary(
    source_dict: dict, keep_list: list) -> dict:
    """Uses the passed ordered list (keep_list) to generate a new 
    dictionary from matching values using the keys from the passed 
    dictionary (source_dict) to generate a new dictionary with 
    corresponding target column numbers in the order of keep_list.
    ex: {(1st matching value from source_dict) "B": "A"}

    Args:
        source_dict (dict): Dictionary of headers and matching columns
        from source Xlsx object.
        keep_list (list): Ordered list of (exact) string values of header
        values that contain the data you'd like to keep.

    Returns:
        dict: Dictionary of source: target letters {"B": "A"}
    """
    return {source_dict[keep_value]: get_column_letter(
        column_number) for column_number, keep_value in enumerate(
            keep_list, 1)}
