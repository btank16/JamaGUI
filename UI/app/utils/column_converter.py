from openpyxl.utils import column_index_from_string

def convert_column_input(column_input):
    """Convert column input (letter or number) to column number"""
    if isinstance(column_input, str):
        # Remove any spaces and convert to uppercase
        column_input = column_input.strip().upper()
        try:
            return column_index_from_string(column_input)
        except:
            raise ValueError(f"Invalid column input: {column_input}")
    else:
        try:
            return int(column_input)
        except:
            raise ValueError(f"Invalid column input: {column_input}") 