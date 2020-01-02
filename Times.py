class Times:
    def __init__(self, year, start_month, start_day, end_month, end_day):
        self.year = year
        self.start_month = start_month
        self.start_day = start_day
        self.end_month = end_month
        self.end_day = end_day


NUMBERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V',
           'W', 'X', 'Y', 'Z']


def get_column_string(index):
    index -= 1
    first_number = index // 26
    second_number = index % 26
    number = ""
    if first_number > 0:
        number += NUMBERS[first_number - 1]
    number += NUMBERS[second_number]
    return number
