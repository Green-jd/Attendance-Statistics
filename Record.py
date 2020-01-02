class Attendance:
    def __init__(self, time, persons):
        self.time = time
        self.persons = persons


class Person:
    def __init__(self, name, number, records):
        self.name = name
        self.number = number
        self.records = records


class Record:
    def __init__(self, time, check_in, check_out):
        self.time = time
        self.check_in = check_in
        self.check_out = check_out
