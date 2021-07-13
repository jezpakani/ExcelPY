from datetime import datetime, date, time, timedelta
from dateutil.parser import parse
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from os import system, name
from colorama import init, Fore
from random import randint

import argparse


class ExcelPY:
    def __init__(self):
        """ Initialize our class and get things ready. """
        self.start_time = datetime.now()  # start timer before first operation
        self.end_time = datetime.now()  # end timer after last operation
        self.execution_time = datetime.now()  # results of our timer

        # filenames for our workbooks (spreadsheets)
        self.fn_alm = 'dump-alm.xlsx'
        self.fn_defect = 'dump-defects.xlsx'
        self.fn_enhancement = 'dump-enhancements.xlsx'
        self.fn_incident = 'dump-incidents.xlsx'
        self.fn_destination = 'capacity-tracker.xlsx'

        # workbook handles
        self.wb_alm = Workbook()
        self.wb_defect = Workbook()
        self.wb_enhancement = Workbook()
        self.wb_incident = Workbook()
        self.wb_destination = Workbook()

        # user supplied options (arguments)
        self.arg_data = False  # generate test data only
        self.arg_check = False  # allows us to run the app to test files without writing to them

        # general application options
        self.date_fields = ['opened', 'planned fix date']
        self.test_data_row_count = 3  # how many rows of test data we will be creating
        self.cells_updated = 0
        self.rows_appended = 0
        self.errors = 0
        self.warnings = 0

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
            self.error(str(e))
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
        """ Populate our input file with random test data. Due to the different number of columns in the files
        and the random nature of this method there will likely be zero matching data between the files.
        """
        self.message('Generating {} rows of unique keyed test data.'.format(self.test_data_row_count))
        if not self.open_workbooks():
            exit()

        # populate our data dump input files
        self.populate_sheet(self.wb_incident, self.wb_incident.active, self.fn_incident, 'Hypercare Incidents', 'INC')
        self.populate_sheet(self.wb_enhancement, self.wb_enhancement.active, self.fn_enhancement,
                            'Hypercare Enhancements', 'ENH')
        self.populate_sheet(self.wb_defect, self.wb_defect.active, self.fn_defect, 'Hypercare Defects', 'DFC')
        self.populate_sheet(self.wb_alm, self.wb_alm.active, self.fn_alm, 'ALM Defects', 'ALM')

        self.message('Completed generating input file')

    def populate_sheet(self, wb, ws, fn, tab, extra):
        """ PARAMETERS:
        wb - workbook
        ws - worksheet (the destination file has multiple sheets)
        fn - the file name to save our changes in
        tab - name of the destination tab (sheet) to write to
        extra - extra info to add to the generated data
        """
        ws.delete_rows(2, ws.max_row + 1)
        self.wb_destination.active = self.wb_destination[tab]
        self.wb_destination.active.delete_rows(2, self.wb_destination.active.max_row + 1)
        used = []

        try:
            # this loop is for the dump file
            for x in range(2, self.test_data_row_count + 2):  # now lets generate some cell data
                for y in range(1, ws.max_column + 1):
                    column_header = ws.cell(row=1, column=y).value
                    if column_header is None:  # no idea why but some sheets not reporting column count correctly.
                        break

                    rand_x = randint(100, 99999)
                    rand_y = randint(100, 99999)
                    buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    while used.count(buffer) != 0:  # disallow duplicate keys
                        rand_x = randint(100, 99999)
                        rand_y = randint(100, 99999)
                        buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    used.append(buffer)  # add our new key to the list

                    if column_header.lower() in self.date_fields:
                        buffer = date.today() + timedelta(days=randint(-10, 35))
                        buffer = buffer.strftime('%Y-%m-%d')
                    # else:
                    #     buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    ws.cell(row=x, column=y).value = buffer

                    if y == 1:  # only write to destination the first time through so keys match
                        self.wb_destination.active.cell(row=x, column=1).value = buffer

                # this loop is for the destination file
                for y in range(2, self.wb_destination.active.max_column + 1):
                    column_header = ws.cell(row=1, column=y).value
                    if column_header is None:  # no idea why but some sheets not reporting column count correctly.
                        break

                    rand_x = randint(100, 99999)
                    rand_y = randint(100, 99999)
                    buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    while used.count(buffer) != 0:  # disallow duplicate keys
                        rand_x = randint(100, 99999)
                        rand_y = randint(100, 99999)
                        buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    used.append(buffer)  # add our new key to the list

                    if column_header.lower() in self.date_fields:
                        buffer = date.today() + timedelta(days=randint(-10, 35))
                        buffer = buffer.strftime('%Y-%m-%d')
                    else:
                        buffer = '[{}] {}:{}'.format(extra, str(rand_x).zfill(5), str(rand_y).zfill(5))

                    self.wb_destination.active.cell(row=x, column=y).value = buffer

            # now set the column widths
            for x in range(1, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(x)].width = 18

            for x in range(1, self.wb_destination.active.max_column + 1):
                self.wb_destination.active.column_dimensions[get_column_letter(x)].width = 18

        except IndexError as e:
            self.error(str(e))
        except Exception as e:
            self.error(str(e))

        wb.save(fn)
        self.wb_destination.save(self.fn_destination)

    def get_execution_time(self):
        """ Display the results of our execution timer. """
        self.execution_time = self.end_time - self.start_time

        print('\n')
        self.message('**[OPERATION COMPLETE]**********************************************************************')
        if self.arg_data:
            self.message(' Execution Time: {} ms'.format(self.execution_time))
            self.message('********************************************************************************************')
        else:
            self.message('   Cell Updates: {}'.format(self.cells_updated))
            self.message(' Cell Additions: {}'.format(self.rows_appended))
            self.message('         Errors: {}'.format(self.errors))
            self.message('       Warnings: {}'.format(self.warnings))
            self.message(' Execution Time: {} ms'.format(self.execution_time))
            self.message('********************************************************************************************')

    def start_timer(self):
        """ Start our timer used to determine execution speed. """
        self.start_time = datetime.now()

    def stop_timer(self):
        """ Stop our timer used to determine execution speed. """
        self.end_time = datetime.now()

    def is_not_used(self):
        """ Useful to get rid of the warning messages about 'self' not being used in a method. """
        pass

    def parse_args(self):
        """ By default without args we will process the files, but we will also give options via
         command line arguments for things like generating test data.
         """
        parser = argparse.ArgumentParser()
        parser.add_argument('-d', '--data', dest='data',
                            help='Generate requested amount of test data.',
                            type=int, nargs='+')
        parser.add_argument('-c', '--check', action='store_true',
                            dest='check', help='Check files without modifying them.',
                            default=False)
        args = parser.parse_args()
        self.arg_data = args.data
        self.arg_check = args.check

        if xc.arg_data:  # did the user request to generate test data?
            choice = input(Fore.YELLOW + 'This option will ' + Fore.RED +
                           '*OVERWRITE ALL FILES* ' + Fore.YELLOW + 'you sure (y/n)? ')
            if choice.upper() == 'Y':
                self.test_data_row_count = int(self.arg_data[0])
                xc.generate_test_data()
            else:
                xc.arg_data = False
        else:
            self.process_dump_files()

    def process_dump_files(self):
        """ Process all the sheets in our workbooks looking for changes. """
        if not self.open_workbooks():
            exit()

        self.message('*****************************************************************************')
        self.message('Only columns that exist in both the dump and destination file will be synced.')
        self.message('They must also match exactly including spelling and capitalization.')
        self.message('*****************************************************************************')

        self.wb_destination.active = self.wb_destination['Hypercare Incidents']
        self.parse_dump_file(self.wb_incident.active, self.wb_destination.active, self.fn_incident)

        # self.wb_destination.active = self.wb_destination['Hypercare Defects']
        # self.parse_dump_file(self.wb_defect.active, self.wb_destination.active, self.fn_defect)
        #
        # self.wb_destination.active = self.wb_destination['Hypercare Enhancements']
        # self.parse_dump_file(self.wb_enhancement.active, self.wb_destination.active, self.fn_enhancement)
        #
        # self.wb_destination.active = self.wb_destination['ALM Defects']
        # self.parse_dump_file(self.wb_alm.active, self.wb_destination.active, self.fn_alm)

    def parse_dump_file(self, ws_dump, ws_dest, fn_dump):
        """ PARAMETERS:
        ws_dump = one of our dump workbooks that we will use as data input
        ws_dest = the name of the worksheet in our output file that should be parsed
        fn_dump = the name of the dump file being parsed
        """
        self.message('BEGIN: [{}] -> [{}]:'.format(fn_dump.upper(), self.fn_destination.upper()), True)
        dump_headers = {}  # column headers from our dump file
        dest_headers = {}  # column headers from our destination file
        comm_headers = {}  # column headers common to both files

        rows_updated = 0
        rows_appended = 0

        # get a list of dump column headers so we can use them for searching
        for x, cell in enumerate(ws_dump[1]):
            dump_headers[cell.value] = x + 1

        # get a list of destination column headers so we can use them for searching
        for x, cell in enumerate(ws_dest[1]):
            dest_headers[cell.value] = x + 1

        # get a list of column headers from both sheets using column locations from destination
        for key1, val1 in dest_headers.items():
            for key2, val2 in dump_headers.items():
                if key1 == key2:
                    comm_headers[key1] = val1
                    break

        # now parse our dump file to arg_check for duplicate 'keys'
        if self.worksheet_has_duplicate_keys(ws_dump, fn_dump):
            return

        # now parse our destination file to arg_check for duplicate 'keys'
        if self.worksheet_has_duplicate_keys(ws_dest, self.fn_destination):
            return

        # let's case-sensitive arg_check our column headers for differences if any.
        s1 = set(dump_headers)
        s2 = set(dest_headers)

        if s1 != s2:
            s1_diff = (s1 - s2)
            s2_diff = (s2 - s1)
            if len(s1_diff) > 0:
                self.warning('{} exclusively contains the following columns: '.format(fn_dump.upper()))
                for x, item in enumerate(s1_diff):
                    self.warning('\t{}. \'{}\''.format(x + 1, str(item)))
            if len(s2_diff) > 0:
                self.warning('{} exclusively contains the following columns: '.format(self.fn_destination.upper()))
                for x, item in enumerate(s2_diff):
                    self.warning('\t{}. \'{}\''.format(x + 1, str(item)))

        # for x in range(2, ws_dump.max_row + 1):
        #     primary_key = ws_dump.cell(row=x, column=1).value
        #     for y in range(1, ws_dump.max_column + 1):
        #         header_name = ws_dump.cell(row=1, column=y).value
        #         results = self.get_cell_value(ws_dest, primary_key, header_name)
        #         if results['cell_found']:  # did we find a cell that matched key and header name?
        #             this = ws_dest.cell(row=results['row'], column=results['col'])
        #             if results['value'] != ws_dump.cell(row=x, column=y).value:  # does the value match dump?
        #                 new_value = ws_dump.cell(row=x, column=y).value
        #                 this.value = new_value
        #                 this.fill = PatternFill(start_color='b2ffff', end_color='b2ffff', fill_type='solid')
        #                 this.font = Font(name='Ubuntu', size=11, color='2e2e2e', bold=False, italic=False)
        #                 self.cells_updated += 1
        #             else:  # there were no changes so reset the cell background
        #                 this.fill = PatternFill(fill_type='none')
        #                 this.font = Font(name='Ubuntu', size=11, color=None, bold=False, italic=False)
        #         else:
        #             if not results['key_found']:  # if the key was not found we need to insert a new row
        #                 this = ws_dest.cell(row=ws_dest.max_row + 1, column=1)
        #                 new_value = ws_dump.cell(row=x, column=y).value
        #                 this.value = new_value
        #                 this.fill = PatternFill(start_color='98fb98', end_color='98fb98', fill_type='solid')
        #                 this.font = Font(name='Ubuntu', size=11, color='2e2e2e', bold=False, italic=False)
        #                 self.rows_appended += 1
        #             else:  # we found the key
        #                 this = ws_dest.cell(row=results['row'], column=1)
        #                 this.fill = PatternFill(fill_type='none')
        #                 this.font = Font(name='Ubuntu', size=11, color=None, bold=False, italic=False)
        #
        # dump_dict = self.parse_worksheet_into_dictionary(ws_dump, dump_headers, dest_headers)
        # dest_dict = self.parse_worksheet_into_dictionary(ws_dest, dest_headers, dest_headers)
        # comb_dict = {**dest_dict, **dump_dict}

        dump_dict = self.parse_worksheet_into_dictionary(ws_dump, comm_headers)
        dest_dict = self.parse_worksheet_into_dictionary(ws_dest, comm_headers)
        comb_dict = {**dest_dict, **dump_dict}

        for a, b in enumerate(comb_dict):  # enumerate dictionary key
            key = b
            dump_row = dump_dict[key]
            for c, d in enumerate(dump_row):  # enumerate dictionary rows
                value = dump_row[d]['value']
                result = self.get_cell_value(ws_dest, key, d)
                if result['key_found']:  # does this key exist in destination
                    this = ws_dest.cell(row=dest_dict[b][d]['row'], column=dest_dict[b][d]['col'])
                    if this.value != value:  # update destination cell
                        self.cells_updated += 1
                        this.value = value
                        this.fill = PatternFill(start_color='00e0e0', end_color='00e0e0', fill_type='solid')
                        this.font = Font(name='Ubuntu', size=11, color='2e2e2e', bold=False, italic=False)
                    else:
                        this.fill = PatternFill(fill_type='none')
                        this.font = Font(name='Ubuntu', size=11, color=None, bold=False, italic=False)
                else:  # key is not present so we are creating a new row
                    self.rows_appended += 1

        # for a, b in enumerate(comb_dict):  # enumerate dictionary key
        #     key = b
        #     row_of_data = dump_dict[key]
        #     for c, d in enumerate(row_of_data):  # enumerate dictionary rows
        #         header = list(row_of_data)[c]
        #         specific_cell_data = row_of_data[d]
        #         value = specific_cell_data['value']
        #         row = specific_cell_data['row']
        #         col = specific_cell_data['col']
        #         this = ws_dest.cell(row=row, column=col)
        #         if this.value != value:  # update destination cell
        #             self.cells_updated += 1
        #             this.value = value
        #             this.fill = PatternFill(start_color='00e0e0', end_color='00e0e0', fill_type='solid')
        #             this.font = Font(name='Ubuntu', size=12, color='2e2e2e', bold=False, italic=False)
        #         else:
        #             this.fill = PatternFill(fill_type='none')
        #             this.font = Font(name='Ubuntu', size=11, color=None, bold=False, italic=False)
        # print('\t\tHeader: \'{}\'\n\t\t\tvalue: \'{}\'\n\t\t\trow: \'{}\'\n\t\t\tcol: \'{}\''.format(header,
        #                                                                                              value, row,
        #                                                                                              col))

        # print('dump_key: {} dump_value {}'.format(key1, val1))

        # dest_dict = self.parse_worksheet_into_dictionary(ws_dest, dest_headers)
        # for key, val in enumerate(dest_dict.values()):
        #     print('dest_key: {} dest_value {}'.format(key, val))

        # for x, row1 in enumerate(ws_dump.values):  # enumerate each row in our dump file
        #     key1 = row1[0]
        #     match = False
        #     for y, row2 in enumerate(ws_dest.values):  # enumerate each row in our destination file
        #         key2 = row2[0]
        #         if key1 == key2:  # arg_check to see if we have matched key fields
        #             match = True
        #             break
        #
        #     if x > 0:  # we skip zero because it is a column header and not a key value
        #         dump_row = x + 1
        #         dest_row = y + 1
        #         if match:  # if we matched keys we need to update cells with new values
        #             for key, value in comm_headers.items():  # we matched keys so now enumerate common headers
        #                 dump_col = dump_headers[key]
        #                 dest_col = dest_headers[key]
        #                 dump_val = ws_dump.cell(dump_row, dump_col).value
        #                 dest_val = ws_dest.cell(dest_row, dest_col).value
        #                 this = ws_dest.cell(dest_row, dest_col)
        #
        #                 # if str(key).lower() in self.date_fields:
        #                 if self.is_date(dump_val):
        #                     dump_val = self.format_date(dump_val)
        #
        #                 if dump_val != dest_val:  # update the cell only if it changed
        #                     this.value = dump_val
        #                     this.fill = PatternFill(start_color='00e0e0', end_color='00e0e0', fill_type='solid')
        #                     this.font = Font(name='Ubuntu', size=12, color='2e2e2e', bold=False, italic=False)
        #                     rows_updated += 1
        #                 else:  # there were no changes so reset the cell background
        #                     this.fill = PatternFill(fill_type='none')
        #                     this.font = Font(name='Ubuntu', size=12, color=None, bold=False, italic=False)
        #
        #                 # if key.lower() in self.date_fields:
        #                 if self.is_date(dest_val):
        #                     cell_date = self.format_date(dest_val)
        #                     today = self.format_date((date.today().strftime('%Y-%m-%d')))
        #                     if cell_date < today:
        #                         this.fill = PatternFill(start_color='ffa8d4', end_color='ffa8d4', fill_type='solid')
        #
        #         else:  # key was not found in destination so we need to append a new row
        #             for key, value in comm_headers.items():  # we matched keys so now enumerate common headers
        #                 dump_col = dump_headers[key]
        #                 dest_col = dest_headers[key]
        #                 dump_val = ws_dump.cell(dump_row, dump_col).value
        #                 this = ws_dest.cell(dest_row + 1, dest_col)
        #
        #                 # if key.lower() in self.date_fields:
        #                 if self.is_date(dump_val):
        #                     dump_val = datetime.strptime(str(dump_val)[0:10], '%Y-%m-%d')
        #                     this.number_format = 'yyyy-mm-dd'
        #                     this.value = dump_val
        #                 else:
        #                     this.value = dump_val
        #
        #                 this.fill = PatternFill(start_color='00e0e0', end_color='00e0e0', fill_type='solid')
        #                 this.font = Font(name='Ubuntu', size=12, color='2e2e2e', bold=False, italic=False)
        #                 this.fill = PatternFill(start_color='00e0e0', end_color='00e0e0', fill_type='solid')
        #                 this.font = Font(name='Ubuntu', size=12, color='2e2e2e', bold=False, italic=False)
        #                 rows_appended += 1

        # save our workbook with all changes
        self.cells_updated += rows_updated
        self.rows_appended += rows_appended
        self.message('END: [{}]  [Updates: {}] [Additions: {}]'.format(fn_dump.upper(), rows_updated, rows_appended))

        # set the active worksheet so it opens on this tab
        if not self.arg_check:
            self.wb_destination.active = self.wb_destination['Hypercare Incidents']
            self.wb_destination.save(self.fn_destination)

    def get_cell_value(self, ws, primary_key, header_name):
        """
        :param ws: the worksheet to search
        :param primary_key: the key field to search
        :param header_name: the column header to search
        :return: dictionary of results
        NOTES: This method does not match values, it only retrieves the failure found at the intersection
               of primary_key,header name which is effectively row,col.
        """
        results = {'cell_found': False, 'key_found': False}

        try:
            for x in range(2, ws.max_row + 1):
                if primary_key == ws.cell(row=x, column=1).value:
                    results['key_found'] = True
                    for y in range(1, ws.max_column + 1):
                        if header_name == ws.cell(row=1, column=y).value:
                            results['value'] = ws.cell(row=x, column=y).value
                            results['cell_found'] = True
                            results['row'] = x
                            results['col'] = y
                            return results
        except Exception as e:
            self.error(str(e))

        return results

    def parse_worksheet_into_dictionary(self, ws, headers):
        """
        :param ws: worksheet to parse
        :param headers: list of common headers in order to match destination columns
        :return: dictionary of data in the proper column order
        """
        self.is_not_used()
        result = {}

        try:
            for x in range(2, ws.max_row + 1):  # enumerate rows
                data = {}
                pkey = ws.cell(row=x, column=1).value
                for key, value in headers.items():
                    buffer = self.get_cell_value(ws, pkey, key)
                    if buffer['cell_found']:
                        value = {'value': buffer['value'], 'row': buffer['row'], 'col': buffer['col']}
                        data[key] = value
                        result[pkey] = data

                    # if value == idx:  # enumerate headers in proper order
                    #     for y in range(1, ws.max_column + 1):  # enumerate ws headers looking for a match
                    #         column = ws.cell(row=1, column=y).value
                    #         if column == key:
                    #             value = {'value': ws.cell(row=x, column=y).value, 'row': x, 'col': y}
                    #             data[column] = value
                    #             break

                        # result[pkey] = data
        except Exception as e:
            self.error(str(e))

        return result

    # def parse_worksheet_into_dictionary(self, ws, headers, common_headers):
    #     self.is_not_used()
    #     result = {}
    #
    #     try:
    #         for x in range(2, ws.max_row + 1):
    #             data = {}
    #             key = ws.cell(row=x, column=1).value
    #             for y in range(1, ws.max_column + 1):
    #                 value = {'value': ws.cell(row=x, column=y).value, 'row': x, 'col': y}
    #                 header_keys = list(headers.keys())
    #                 header_key = header_keys[y - 1]
    #                 if header_key in list(common_headers.keys()):
    #                     data[header_key] = value
    #
    #             result[key] = data
    #     except Exception as e:
    #         self.error(str(e))
    #
    #     return result

    def worksheet_has_duplicate_keys(self, ws, fn):
        self.is_not_used()
        results = {}

        for x in ws.iter_rows(2, ws.max_row, values_only=True):  # enumerate our worksheet keys
            key = x[0]
            if key in results:  # see if key is already in the dictionary
                results[key] = results[key] + 1  # if yes then increment found counter
            else:
                results[key] = 1  # key wasn't in the dictionary so add it now

        for key, value in list(results.items()):  # enumerate our keys
            if results[key] == 1:  # if value > 1 then it is a duplicate key
                del results[key]  # not a duplicate so remove from dictionary
            else:
                results[key] = 'occurrences: ' + str(value)

        if len(results.keys()) > 0:
            self.error(
                '[{}] ({}) contains the following duplicate keys in the first column:'.format(fn.upper(), ws.title))
            self.error(str(results))
            return True
        else:
            return False

    # def find_value_in_worksheet(self, needle, haystack, column='', key=''):
    #     """PARAMETERS:
    #     needle - the text to search for
    #     haystack - the worksheet to search in
    #     column - (optional) search only this column
    #     key - (optional) search on the row this key is found in
    #     """
    #     self.is_not_used()
    #     result = {'found': False, 'coordinates': '', 'msg': ''}
    #
    #     for col in range(1, haystack.max_column + 1):
    #         for row in range(1, haystack.max_row + 1):
    #             value = haystack.cell(row, col).value
    #             if value == needle:
    #                 cell = haystack.cell(row, col)
    #                 coordinates = cell.column_letter + str(cell.row)
    #                 result['code'] = 0
    #                 result['coordinates'] = coordinates
    #                 result['row'] = row
    #                 result['col'] = col
    #                 result['msg'] = ''
    #                 return result
    #
    #     return result

    def days_between(self, d1, d2):
        self.is_not_used()
        try:
            d1 = self.format_date(d1)
            d2 = self.format_date(d2)
            d1 = datetime.strptime(d1, '%Y-%m-%d')
            d2 = datetime.strptime(d2, '%Y-%m-%d')
            return abs((d2 - d1).days)
        except Exception as e:
            self.error(str(e))

    def is_date(self, string, fuzzy=False):
        self.is_not_used()
        try:
            parse(string, fuzzy=fuzzy)
            return True

        except ValueError:
            return False

    def format_date(self, date_val):
        try:
            d = date.fromisoformat(date_val[0:10])
            return d.strftime('%Y-%m-%d')
        except Exception as e:
            self.error((str(e)))

    def message(self, value='', line_before=False):
        """ Format general messages including attributes. """
        self.is_not_used()
        if line_before:
            print('\n')
        print(Fore.GREEN + '+++ ' + value)

    def error(self, value='', line_before=False):
        """ Format error messages including attributes. """
        self.errors += 1
        if line_before:
            print('\n')
        print(Fore.RED + '!!! ' + value)

    def warning(self, value='', line_before=False):
        """ Format warning messages including attributes. """
        self.warnings += 1
        if line_before:
            print('\n')
        print(Fore.YELLOW + '--- ' + value)


def clear_screen():
    """ Clear the screen taking into account operating system. """
    if name == "nt":
        system('cls')
    else:
        system('clear')


if __name__ == '__main__':
    clear_screen()
    init(autoreset=True)
    xc = ExcelPY()
    xc.start_timer()
    xc.parse_args()
    xc.stop_timer()
    xc.get_execution_time()
    del xc
    exit(0)
