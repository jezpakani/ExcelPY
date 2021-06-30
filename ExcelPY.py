import datetime
from openpyxl import load_workbook, Workbook
from os import system, name
from colorama import init, Fore


# from random import randint
# from datetime import timedelta, date


class ExcelPY:
    def __init__(self):
        """ Initialize our class and get things ready. """
        self.start_time = datetime.datetime.now()  # start timer before first operation
        self.end_time = datetime.datetime.now()  # end timer after last operation
        self.execution_time = datetime.datetime.now()  # results of our timer

        # filenames for our workbooks (spreadsheets)
        self.fn_alm = 'dump-alm.xlsx'
        self.fn_defect = 'dump-defects.xlsx'
        self.fn_enhancement = 'dump-enhancements.xlsx'
        self.fn_incident = 'dump-incidents.xlsx'
        self.fn_destination = 'cardinal-health-data-tracking.xlsx'

        # workbook handles
        self.wb_alm = Workbook()
        self.wb_defect = Workbook()
        self.wb_enhancement = Workbook()
        self.wb_incident = Workbook()
        self.wb_destination = Workbook()

        self.test_data_row_count = 500  # how many rows of test data we will be creating

    def __del__(self):
        """ Perform cleanup operations when we destroy our class instance. """
        self.close_files()

    def open_workbooks(self):
        """ Open files needed to perform our processes. """
        try:
            self.wb_alm = load_workbook(self.fn_alm)
            self.wb_defect = load_workbook(self.fn_defect)
            self.wb_enhancement = load_workbook(self.fn_enhancement)
            self.wb_incident = load_workbook(self.fn_incident)
            self.wb_destination = load_workbook(self.fn_destination)
        except Exception as e:
            error(str(e))
            return False

        return True

    def close_files(self):
        """ Close all of the files that we used. """
        self.wb_alm.close()
        self.wb_defect.close()
        self.wb_enhancement.close()
        self.wb_incident.close()
        self.wb_destination.close()

    def generate_test_data(self):
        """ Populate our input file with random test data """
        message('Generating input file with random values')
        if not self.open_workbooks():
            exit()

        # populate our data dump input files
        self.populate_sheet(self.wb_alm, self.wb_alm.active, self.fn_alm, 'alm')
        self.populate_sheet(self.wb_defect, self.wb_defect.active, self.fn_defect, 'dfc')
        self.populate_sheet(self.wb_incident, self.wb_incident.active, self.fn_incident, 'inc')
        self.populate_sheet(self.wb_enhancement, self.wb_enhancement.active, self.fn_enhancement, 'enh')

        # now populate our main reporting file with data that perfectly matches our dumps
        self.populate_sheet(self.wb_destination, self.wb_destination['ALM Defects'],
                            self.fn_destination, 'alm')
        self.populate_sheet(self.wb_destination, self.wb_destination['Hypercare Defects'],
                            self.fn_destination, 'dfc')
        self.populate_sheet(self.wb_destination, self.wb_destination['Hypercare Incidents'],
                            self.fn_destination, 'inc')
        self.populate_sheet(self.wb_destination, self.wb_destination['Hypercare Enhancements'],
                            self.fn_destination, 'enh')

        message('Completed generating input file')

    def populate_sheet(self, wb, ws, fn, tag):
        ws.delete_rows(2, ws.max_row + 1)

        for x in range(2, self.test_data_row_count + 1):  # now lets generate some cell data
            for y in range(1, ws.max_column + 1):  # worksheet columns are not zero-based so add 1
                ws.cell(row=x, column=y).value = '[{}]:{}:{}'.format(tag, x, y)

        wb.save(fn)
        # ws_hdr = list()
        # get a list of column headers so we can use them for searching
        # for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=1, values_only=True):
        #     ws_hdr.append(row)

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
    print(Fore.GREEN + '+++ ' + value + ' +++')


def error(value=''):
    """ Format error messages including attributes. """
    print(Fore.RED + '!!! ' + value + ' !!!')


def warning(value=''):
    """ Format warning messages including attributes. """
    print(Fore.YELLOW + '!!! ' + value + ' !!!')


if __name__ == '__main__':
    clear_screen()
    init(autoreset=True)
    message('Initiating Process')
    xc = ExcelPY()
    xc.start_timer()
    xc.generate_test_data()
    xc.stop_timer()
    xc.get_execution_time()
    del xc
    exit(0)
