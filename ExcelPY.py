import datetime
import argparse

from openpyxl import load_workbook, Workbook, utils
from os import system, name
from colorama import init, Fore, Style
from random import randint


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

        # user supplied options (arguments)
        self.arg_data = False  # generate test data only

        # general application options
        self.test_data_row_count = 50  # how many rows of test data we will be creating

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
        """ Populate our input file with random test data. We will default so that the data in the dump files
        matches that in the destination file. This allows us to start from a known point and we can then change
        data anywhere to test the application.
        """
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
        """ PARAMETERS:
        wb - workbook
        ws - worksheet (the destination file has multiple sheets)
        fn - the file name to save our changes in
        tag - extra info to add to the generated data
        """
        ws.delete_rows(2, ws.max_row + 1)

        try:
            for x in range(2, self.test_data_row_count + 1):  # now lets generate some cell data
                for y in range(1, ws.max_column + 1):  # worksheet columns are not zero-based so add 1
                    rand_x = randint(x, 99)
                    rand_y = randint(y, 99)
                    ws.cell(row=x, column=y).value = '[{}] {}:{}'.format(tag, rand_x, rand_y)
        except Exception as e:
            print(str(e))

        wb.save(fn)

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

    def parse_args(self):
        """ By default without args we will process the files, but we will also give options via
         command line arguments for things like generating test data.
         """
        parser = argparse.ArgumentParser()
        parser.add_argument('-d', '--data', action='store_true',
                            dest='data', help='Generate test data (overwrites all files)',
                            default=False)
        args = parser.parse_args()
        self.arg_data = args.data

        if xc.arg_data:  # did the user request to generate test data?
            choice = input(Fore.YELLOW + 'This option will ' + Fore.RED +
                           '*OVERWRITE ALL FILES* ' + Fore.YELLOW + 'you sure (y/n)?')
            if choice.upper() == 'Y':
                xc.generate_test_data()
            else:
                xc.arg_data = False
        else:
            self.process_dump_files()

    def process_dump_files(self):
        """ Process all the sheets in our workbooks looking for changes. """
        message('Initializing workbook processing')
        if not self.open_workbooks():
            exit()

        # self.parse_dump_file(self.wb_incident.active, self.wb_destination['Hypercare Incidents'], self.fn_incident)
        self.wb_destination.active = self.wb_destination['Hypercare Incidents']
        self.parse_dump_file(self.wb_incident.active, self.wb_destination.active, self.fn_incident)
        # self.parse_dump_file(self.wb_alm, 'ALM Defects', self.fn_alm)
        # self.parse_dump_file(self.wb_defect, 'Hypercare Defects', self.fn_defect)
        # self.parse_dump_file(self.wb_enhancement, 'Hypercare Enhancements', self.fn_enhancement)

    def parse_dump_file(self, ws_dump, ws_dest, fn_dump):
        """ PARAMETERS:
        ws_dump = one of our dump workbooks that we will use as data input
        ws_dest = the name of the worksheet in our output file that should be parsed
        fn_dump = the name of the dump file being parsed
        """
        message('Processing dump file [{}] -> [{}]:'.format(fn_dump.upper(), self.fn_destination.upper()))
        dump_headers = {}  # column headers from our dump file
        dest_headers = {}  # column headers from our destination file
        comm_headers = {}  # column headers common to both files

        # get a list of dump column headers so we can use them for searching
        for x, cell in enumerate(ws_dump[1]):
            dump_headers[cell.value] = x + 1

        # get a list of destination column headers so we can use them for searching
        for x, cell in enumerate(ws_dest[1]):
            dest_headers[cell.value] = x + 1

        # get a list of column headers that are common to both sheets
        for key1, val1 in dump_headers.items():
            for key2, val2 in dest_headers.items():
                if key1 == key2:
                    comm_headers[key2] = val2

        # let's case-sensitive check our column headers for differences if any.
        s1 = set(dump_headers)
        s2 = set(dest_headers)

        if s1 != s2:
            s1_diff = (s1 - s2)
            s2_diff = (s2 - s1)
            warning('The dump file and destination file have different column headers.')
            warning('{} exclusively contains {}: '.format(fn_dump.upper(), s1_diff))
            warning('{} exclusively contains {}: '.format(self.fn_destination.upper(), s2_diff))
            warning('Be sure to check for misspellings, capitalization, and spaces.')
            warning('Unmatched column headers regardless of reason WILL NOT be updated.')

        for x, row1 in enumerate(ws_dump.values):
            key1 = row1[0]
            match = False
            for y, row2 in enumerate(ws_dest.values):
                key2 = row2[0]
                if key1 == key2:
                    match = True
                    break

            if x > 0:  # zero is our header rows
                if not match:
                    for key, value in comm_headers.items():
                        print('key: {}, value: {}'.format(key, value))
                else:
                    pass
                    # print('Update: {}'.format(key1))

            # print('Key: {} Found: {}'.format(key1, match))

        # # save our workbook with all changes
        # self.wb_destination.save(self.fn_destination)

    def find_value_in_destination_column(self, sheet, haystack, needle, dump_fn):
        """PARAMETERS:
        sheet - the tab name (worksheet) we will be searching
        haystack - the column to search
        needle - the value to find
        fn -

        NOTES:
            This method will search our destination excel file for 'needle' in 'haystack' on 'sheet'.
        """
        self.is_not_used()
        wb_dest = self.wb_destination
        wb_dest.active = wb_dest[sheet]
        column_found = False
        result = {'code': 0, 'coordinates': '', 'msg': ''}

        # loop through all the column headers looking for our match to haystack
        for col in range(1, wb_dest.active.max_column + 1):
            if column_found:  # break if we already searched the requested column
                break
            column = wb_dest.active.cell(1, col).value
            if column == haystack:  # we found the correct column to search so now let's search its rows
                column_found = True
                # now enumerate through all the rows in that column
                for row in range(2, wb_dest.active.max_row + 1):
                    value = wb_dest.active.cell(row, col).value
                    if value == needle:
                        cell = wb_dest.active.cell(row, col)
                        coordinates = cell.column_letter + str(cell.row)
                        result['code'] = 0
                        result['coordinates'] = coordinates
                        result['row'] = row
                        result['col'] = col
                        result['msg'] = '[{}] - Success: \'{}\' in column \'{}\' at \'{}\'.'.format(dump_fn, needle,
                                                                                                    haystack,
                                                                                                    coordinates)
                        return result

        # we did not find the needle in the haystack so return nothing
        if column_found:
            # we found our column but not our value meaning this should be a new row
            result['code'] = 1
            result['msg'] = '[{}] - Failure: value \'{}\' was not found in column \'{}\'.'.format(dump_fn, needle,
                                                                                                  haystack)
        else:
            # we were not able to find a column header in destination that matched
            result['code'] = 2
            result['msg'] += '[{}] - Failure: column \'{}\' was not found.'.format(dump_fn, haystack)
        return result

    def find_value_in_worksheet(self, ws, needle):
        """PARAMETERS:
        sheet - the tab name (worksheet) we will be searching
        haystack - the column to search
        needle - the value to find

        NOTES:
            This method will search our destination excel file for 'needle' in 'haystack' on 'sheet'.
        """
        self.is_not_used()
        result = {'code': -1, 'coordinates': '', 'msg': ''}

        for col in range(1, ws.max_column + 1):
            for row in range(1, ws.max_row + 1):
                value = ws.cell(row, col).value
                if value == needle:
                    cell = ws.cell(row, col)
                    coordinates = cell.column_letter + str(cell.row)
                    result['code'] = 0
                    result['coordinates'] = coordinates
                    result['row'] = row
                    result['col'] = col
                    result['msg'] = ''
                    return result

        return result


def clear_screen():
    """ Clear the screen taking into account operating system. """
    if name == "nt":
        system('cls')
    else:
        system('clear')


def message(value=''):
    """ Format general messages including attributes. """
    print(Fore.GREEN + '+++ ' + value)


def error(value=''):
    """ Format error messages including attributes. """
    print(Fore.RED + '!!! ' + value)


def warning(value=''):
    """ Format warning messages including attributes. """
    print(Fore.YELLOW + '--- ' + value)


if __name__ == '__main__':
    clear_screen()
    init(autoreset=True)
    message('Initiating Process')
    xc = ExcelPY()
    xc.start_timer()
    xc.parse_args()
    xc.stop_timer()
    xc.get_execution_time()
    del xc
    exit(0)
