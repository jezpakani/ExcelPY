import datetime
from openpyxl import load_workbook, Workbook
from os import system, name
from colorama import init, Fore
from random import randint
from datetime import timedelta, date


class ExcelPY:
    def __init__(self, infile_name, outfile_name):
        """ Initialize our class and get things ready. """
        self.start_time = datetime.datetime.now()  # start timer before first operation
        self.end_time = datetime.datetime.now()  # end timer after last operation
        self.execution_time = datetime.datetime.now()  # results of our timer
        self.infile_name = infile_name  # name of input file (data dump)
        self.outfile_name = outfile_name  # name of output file (target excel spreadsheet)
        self.infile = Workbook()  # file handle for input file
        self.outfile = Workbook()  # file handle for output file
        self.test_data_row_count = 1000  # how many rows of test data we will be creating

        if not self.open_files():
            exit()

    def __del__(self):
        """ Perform cleanup operations when we destroy our class instance. """
        self.close_files()

    def open_files(self):
        """ Open files needed to perform our processes. """
        try:
            self.infile = load_workbook(self.infile_name)
        except Exception as e:
            error(str(e))
            return False

        try:
            self.outfile = load_workbook(self.outfile_name)
        except Exception as e:
            self.close_files()
            error(str(e))
            return False

        return True

    def close_files(self):
        """ Close all of the files that we used. """
        self.infile.close()
        self.outfile.close()

    def generate_data_dump(self):
        """ Populate our input file with random test data """
        # TODO: Add logic to this method to add created data to the correct tab in our output file.
        message('Generating input file with random values')
        self.is_not_used()
        column_headers = ('Type Number', 'Opened', 'Short Description', 'Impact', 'Priority',
                          'Severity', 'Status', 'Opened By', 'Assigned To', 'Incident Status',
                          'SLA Due', 'Estimated Hours', 'Assignment Group')
        data_types = ('INC', 'DFCT', 'ENHC', 'ALM')
        levels = ('1 - Low', '2 - Medium', '3 - High')
        status = ('Open', 'Closed')
        people = ('Pete Skeebo', 'Sam Sausagehead', 'Mrs. Buttersworth', 'Jax Vonhapsburg')

        sheet = self.infile.active  # grab the first sheet
        sheet.title = 'Dump'
        sheet.delete_rows(1, sheet.max_row + 1)  # now delete all the rows as we are starting from scratch

        for x in range(0, len(column_headers)):  # create our column headers
            sheet.cell(row=1, column=x + 1).value = column_headers[x]

        for x in range(2, self.test_data_row_count):  # now lets generate random cell data
            sheet.cell(row=x, column=1).value = data_types[randint(0, len(data_types) - 1)] + str(randint(10000, 99999))
            sheet.cell(row=x, column=2).value = date.today() + timedelta(days=randint(1, 10))
            sheet.cell(row=x, column=3).value = 'Short Description #' + str(x)
            sheet.cell(row=x, column=4).value = levels[randint(0, len(levels) - 1)]
            sheet.cell(row=x, column=5).value = levels[randint(0, len(levels) - 1)]
            sheet.cell(row=x, column=6).value = levels[randint(0, len(levels) - 1)]
            sheet.cell(row=x, column=7).value = status[randint(0, len(status) - 1)]
            sheet.cell(row=x, column=8).value = people[randint(0, len(people) - 1)]
            sheet.cell(row=x, column=9).value = people[randint(0, len(people) - 1)]
            sheet.cell(row=x, column=10).value = status[randint(0, len(status) - 1)]
            sheet.cell(row=x, column=11).value = date.today() + timedelta(days=randint(1, 10))
            sheet.cell(row=x, column=12).value = randint(1, 10)
            sheet.cell(row=x, column=13).value = 'Assignment Group #' + str(randint(1, 10))

        self.infile.save(self.infile_name)
        message('Completed generating input file')

    def get_execution_time(self):
        """ Display the results of our execution timer. """
        self.execution_time = self.end_time - self.start_time
        message('Total execution time was {} ms.'.format(self.execution_time))

    def start_timer(self):
        """ Start our timer used to determine execution speed. """
        self.start_time = datetime.datetime.now()

    def stop_timer(self):
        """ Stop our timer used to determine execution speed. """
        self.end_time = datetime.datetime.now()

    def is_not_used(self):
        """ Useful to get rid of the warning messages about 'self' not being used in a method. """
        pass


def clear_screen():
    """ Clear the screen taking into account operating system. """
    if name == "nt":
        system('cls')
    else:
        system('clear')


def message(value=''):
    """ Format general messages including attributes. """
    print(Fore.YELLOW + '+++ ' + value + ' +++')


def error(value=''):
    """ Format error messages including attributes. """
    print(Fore.RED + '+++ ' + value + ' +++')


if __name__ == '__main__':
    clear_screen()
    init(autoreset=True)
    message('Initiating Process')
    xc = ExcelPY('input_data.xlsx', 'output_data.xlsx')
    xc.start_timer()
    xc.generate_data_dump()
    xc.stop_timer()
    xc.get_execution_time()
    del xc
    exit(0)
