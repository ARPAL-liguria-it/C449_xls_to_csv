import unittest
import unittest.mock as mock

from converter import converter
import pandas as pd


class TestReadCsv(unittest.TestCase):
    def test_read_content(self):
        data = converter.read_content('./data/foglio_calcolo_diluizioni.xlsx',
                                      sheet='perCalcoliDiluizione',
                                      start_row=102,
                                      end_row=158,
                                      start_column='A',
                                      end_column='B')

        self.assertIsInstance(data, pd.core.frame.DataFrame)
        self.assertEqual(data.shape[0], 55)
        self.assertEqual(data.iloc[[0],[0]].to_string(header=False, index=False),
                         'Codice campione')
        self.assertEqual(data.iloc[[54],[1]].to_string(header=False, index=False),
                         ' 13.798632')

    def test_convert_to_csv(self):
        test_df = converter.read_content('./data/foglio_calcolo_diluizioni.xlsx',
                                         sheet='perCalcoliDiluizione',
                                         start_row=102,
                                         end_row=158,
                                         start_column='A',
                                         end_column='B')

        with mock.patch.object(test_df, 'to_csv') as to_csv_mock:
            converter.convert_to_csv(test_df, 'test.csv')
            to_csv_mock.assert_called_with('test.csv',
                                           header=False,
                                           index=False,
                                           quoting=4,
                                           sep=';',
                                           encoding="utf-8")

if __name__ == '__main__':
    unittest.main()
