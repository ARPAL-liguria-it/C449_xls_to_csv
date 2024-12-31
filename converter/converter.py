from os import listdir
import re
import csv
from tkinter import filedialog as fd
from tkinter import messagebox as mb
import pandas as pd
from decimal import Decimal, getcontext

def list_xls(dirpath):
    """
    A function to list Excel files (.xsl and .xlsx) in a directory
    :param dirpath: str
        the path of the directory for which Excel files should be listed.
    :return: list
        filenames of xls and xlsx files within the specified directory.
    """
    files = listdir(dirpath)
    # find xls and xlsx files and avoid hidden files
    files_xls = [i for i in files if (i.endswith('.xls') or i.endswith('.xlsx') and not i.startswith('.'))]

    return files_xls

def read_content(filepath, sheet, start_row, end_row, start_column, end_column):
    """
    A function for reading a delimited box area in an Excel file

    :param filepath: str
        the full path of the file to be read.
    :param sheet: str
        the name of the sheet to be read.
    :param start_row: int
        the first row index to be read.
    :param end_row: int
        the last row index to be read.
    :param start_column: str
        the first column to be read.
    :param end_column: str
        the last column to be read.
    :return: DataFrame
        a pandas DataFrame with the content of the specified area in the selected sheet and file.
    """
    first_row = start_row - 1
    n_rows = end_row - start_row

    data = pd.read_excel(filepath,
                         sheet_name=sheet,
                         usecols=start_column + ':' + end_column,
                         skiprows=first_row,
                         nrows=n_rows,
                         engine='openpyxl')

    return data


def convert_to_csv(values, filepath):
    """
    A function for converting a pandas DataFrame to a csv file
    :param values: DataFrame
        a pandas DataFrame obtained by the pandas read_excel function.
    :param filepath: str
        the path for the csv file to be saved.
    :return: csv file
        A csv file, containing the DataFrame specified in the values parameter,
        will be saved at the specified filepath.
    """
    values.to_csv(filepath,
                  header=False,
                  index=False,
                  quoting=csv.QUOTE_STRINGS,
                  sep =';',
                  encoding="utf-8")

def clean_names(filenames):
    """
    A function for removing non-alphanumerical characters from filenames
    :param filenames: list
        A list of filenames, each provided as a string.
    :return: list
        A list of filenames without non-alphanumerical characters and file extension.
    """
    cleaned_names = [re.sub('[^0-9a-zA-Z_]+', '', i.rsplit('.', 1)[0]) for i in filenames]

    return cleaned_names

def round_to_sig_figs(x, sigfigs):
    """
    Round a number off to a given number of significant figures.
    from https://rowannicholls.github.io/python/data/rounding_off.html
    :param x: float
        value to be converted
    :param sigfigs: int
        number of desired significant figures
    """
    # A special case is if the number is 0
    if float(x) == 0.0:
        output = '0' * sigfigs
        # Insert a decimal point
        output = output[0] + '.' + output[1:]

        return output

    # A special case is if the number is 'nan'
    if pd.isna(x):

        return 'None'

    # Set the precision
    getcontext().prec = sigfigs
    rounded = Decimal(x)
    # The precision only kicks in when an arithmetic operation is performed
    rounded = rounded * 1
    # Remove scientific notation
    # (if the order of magnitude of the number is larger than the number of
    # significant figures, the number will have been converted to scientific
    # notation)
    rounded = float(rounded)
    # Convert to string
    output = str(rounded)

    # Count the number of significant figures
    sigfigs_now = len(output.replace('-', '').replace('.', '').lstrip('0'))

    # Remove trailing zero if one exists and it is not necessary
    if (sigfigs_now > sigfigs) and (output.endswith('.0')):
        output = output.removesuffix('.0')

    # Add trailing zeroes if necessary
    if sigfigs_now < sigfigs:
        discrepancy = sigfigs - sigfigs_now
        # Append a decimal point if necessary
        if '.' not in output:
            output = output + '.'
        # Add trailing zeroes
        output = output + '0' * discrepancy

    return output

def apply_sigfig(data, rows, column, sigfig):
    """
    A function for formatting numbers in a column of a Pandas data.frame
    with the desired number of significant digits
    :param data: DataFrame
        a pandas DataFrame with named columns.
    :param rows: list
        the row indexes with the numbers to be converted.
    :param column: int
        the number of the column containing the values to be formatted.
    :param sigfig: int
        the number of the desired significant figures.
    :return: DataFrame
        a pandas DataFrame.
    """
    data.iloc[rows, column] = data.iloc[rows, column].apply(lambda x: round_to_sig_figs(x, sigfig))

    return data


def main():
    """
    A function for reading a defined portion of all the Excel files (xls and xlsx files)
    found in a folder specified by the user and saving the content in csv files at
    a user specified folder.
    """

    # sheet name to be found in the Excel files
    mysheet = 'perCalcoliDiluizione'
    # rows id with numeric values to be rounded
    num_rows = list(range(3, 47)) + list(range(48, 55))
    # column id with numeric values to be rounded
    num_col = 1
    # number of significant digits required for rounding
    required_sigfig = 2

    # select the folder
    path_xls = fd.askdirectory(title="Seleziona la cartella contenente i file Excel da convertire",
                               mustexist=True)
    if path_xls:
        # read the filenames in the selected folder.
        files_xls = list_xls(path_xls)
        fullpath_xls = [path_xls + "/" + i for i in files_xls]
        files_n = len(files_xls)

        # dialog box with the number of xls and xlsx files found in the selected folder
        if files_n >= 1:
            res = mb.askquestion('File Excel trovati',
                                 f'Ho trovato {files_n} file Excel.\n'
                                 f'Vuoi procedere alla conversione in file csv?')

            # select the folder for the converted csv files
            if res == 'yes':
                path_csv = fd.askdirectory(title="Seleziona la cartella dove salvare i file csv",
                                           mustexist=True)

                if path_csv:
                    # removing non-alphanumerical characters (except underscore) from filenames
                    files_csv = clean_names(files_xls)
                    fullpath_csv = [path_csv + "/" + i + ".csv" for i in files_csv]

                    content_csv = []
                    missing_sheet = 0
                    for i in range(files_n):
                        try:
                            # read the content of each Excel file
                            content_csv = read_content(fullpath_xls[i],
                                        sheet=mysheet,
                                        start_row=102,
                                        end_row=158,
                                        start_column='A',
                                        end_column='B')

                            # format the numbers with the required significant figures
                            apply_sigfig(content_csv,
                                         rows = num_rows,
                                         column = num_col,
                                         sigfig = required_sigfig)

                            # convert the content to a csv file
                            convert_to_csv(content_csv,
                                           fullpath_csv[i])
                        except ValueError:
                            print('Scheda non trovata')
                            missing_sheet += 1
                        except IndexError:
                            pass

                    # recap message
                    if missing_sheet != 0:
                        warning_msg = f'\n In {missing_sheet} file mancava la scheda {mysheet}.'
                    else:
                        warning_msg = ''

                    mb.showinfo('File convertiti',
                                f'Sono stati convertiti {files_n - missing_sheet} file Excel.\n'
                                f'I file csv sono stati salvati nella cartella {path_csv}.\n'
                                f'\n'
                                f'{warning_msg}')

                # no folder selected, exiting the script
                else:
                    mb.showinfo('Operazione annullata',
                                'Nessuna cartella selezionata.\n'
                                'Chiusura dell\'applicazione')

        else:
            # no Excel file found
            mb.showinfo('File Excel non trovati',
                        'Nessun file Excel nella cartella selezionata.')

    else:
        # no folder selected, exiting the script
        mb.showinfo('Operazione annullata',
                    'Nessuna cartella selezionata.\n'
                    'Chiusura dell\'applicazione')


if __name__ == "__main__":
    main()
