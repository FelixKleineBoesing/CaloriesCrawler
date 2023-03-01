def convert_column_number_to_excel_column_name(column_number: int) -> str:
    dividend = column_number
    column_name = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        column_name = chr(65 + modulo) + column_name
        dividend = (dividend - modulo) // 26
    return column_name


def get_sum_formulas(number_columns: int):
    starts = [17, 18, 19, 20]
    steps = 7
    single_columns = [[start + i * steps for i in range(number_columns)] for start in starts]
    inner_values = [[
        f'{convert_column_number_to_excel_column_name(column)}2' for column in column_list
    ] for column_list in single_columns]
    inner_values = ["; ".join(value_list) for value_list in inner_values]

    for column in inner_values:
        print(f'=SUMME({column})')


if __name__ == "__main__":
    get_sum_formulas(30)