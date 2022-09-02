import openpyxl

class Employer:
    def __init__(self, name, surname, stage):
        self.name=name
        self.surname=surname
        self.stage=stage

    def get_initials(self):
        self.initials=self.name[0].upper()+self.surname[0].upper()
        return self.initials

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.name