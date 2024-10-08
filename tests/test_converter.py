import unittest
from unittest.mock import patch as patch
import pandas as pd
from converter import converter

class TestConvertCsv(unittest.TestCase):
    def setUp(self):
        self.dir = './data'
        self.filename = './data/foglio_calcolo_diluizioni.xlsx'
        self.data = converter.read_content(self.filename,
                                           sheet='perCalcoliDiluizione',
                                           start_row=102,
                                           end_row=158,
                                           start_column='A',
                                           end_column='B')
        self.badnames = ['Cus ter145à-!.csv', 'S--_à3.txt', '12Dfg_-ò.xlsx']

    def tearDown(self):
        del self.dir
        del self.filename
        del self.data
        del self.badnames

    def test_listxls(self):
        self.assertEqual(len(converter.list_xls(self.dir)), 1)
        self.assertEqual(converter.list_xls(self.dir), ['foglio_calcolo_diluizioni.xlsx'])

    def test_read_as_dataframe(self):
        """Testing that the Excel file is read as pandas DataFrame"""
        self.assertIsInstance(self.data, pd.core.frame.DataFrame)

    def test_dataframe_size(self):
        """Checking the size of the read content"""
        self.assertEqual(self.data.shape[0], 55)

    def test_dataframe_content(self):
        """Testing the content at the corner of the DataFrame"""
        self.assertEqual(self.data.iloc[[0],[0]].to_string(header=False, index=False),
                         'Codice campione')
        self.assertEqual(self.data.iloc[[54],[1]].to_string(header=False, index=False),
                         ' 13.798632')

    def test_clean_names(self):
        """Testing the substitution of non-alphanumerical characters from filenames"""
        self.assertEqual(converter.clean_names(self.badnames), ['Custer145', 'S_3', '12Dfg_'])

    def test_converter_call(self):
        """Testing the call to the csv converter"""
        with patch.object(self.data, 'to_csv') as to_csv_mock:
            converter.convert_to_csv(self.data, 'test.csv')
            to_csv_mock.assert_called_with('test.csv',
                                           header=False,
                                           index=False,
                                           quoting=4,
                                           sep=';',
                                           encoding="utf-8")

if __name__ == '__main__':
    unittest.main()
