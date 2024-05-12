'''import unittest
import pandas as pd
from g_ratio import calculate_g_ratio  

class TestCalculateGRatio(unittest.TestCase):

    def setUp(self):
        # Setup test data with complete column set
        self.data = {
            'CTL1_Ax': [0.5, 0.6, 0.7],
            'CTL1_My': [0.2, 0.3, 0.4],
            'CTL2_Ax': [0.8, 0.9, 1.0],
            'CTL2_My': [0.4, 0.5, 0.6]
        }
        self.df = pd.DataFrame(self.data)
        self.experiments = [
            ('CTL1_Ax', 'CTL1_My'), 
            ('CTL2_Ax', 'CTL2_My')
        ]

    def test_calculate_g_ratio(self):
        # Run the calculate_g_ratio function
        result_df = calculate_g_ratio(self.df.copy(), self.experiments)

        # Verify the existence and correctness of G-ratio calculations
        expected_g_ratios_ctl1 = [0.71, 0.67, 0.64]  
        expected_g_ratios_ctl2 = [0.67, 0.64, 0.62]
        pd.testing.assert_series_equal(result_df['G-Ratio_CTL1'], pd.Series(expected_g_ratios_ctl1, name='G-Ratio_CTL1'))
        pd.testing.assert_series_equal(result_df['G-Ratio_CTL2'], pd.Series(expected_g_ratios_ctl2, name='G-Ratio_CTL2'))

    def test_ensure_columns_exist(self):
        # Ensure all required columns are present before calling the function
        for ax_col, my_col in self.experiments:
            self.assertIn(ax_col, self.df.columns)
            self.assertIn(my_col, self.df.columns)

        # Optionally, run the calculation function after confirming column presence
        result_df = calculate_g_ratio(self.df.copy(), self.experiments)
        self.assertIn('G-Ratio_CTL1', result_df.columns)
        self.assertIn('G-Ratio_CTL2', result_df.columns)

'''

import unittest
import pandas as pd
from g_ratio import calculate_g_ratio  

class TestCalculateGRatio(unittest.TestCase):

    def setUp(self):
        # Setup test data with your provided column set
        self.data = {
            'CTL1_Ax': [0.79, 0.85, 0.66],
            'CTL1_My': [0.22, 0.33, 0.02],
            'CTL2_Ax': [0.14, 0.54, 0.63],
            'CTL2_My': [0.23, 0.18, 0.20]
        }
        self.df = pd.DataFrame(self.data)
        self.experiments = [
            ('CTL1_Ax', 'CTL1_My'), 
            ('CTL2_Ax', 'CTL2_My')
        ]

    def test_calculate_g_ratio(self):
        # Run the calculate_g_ratio function
        result_df = calculate_g_ratio(self.df.copy(), self.experiments)

        # Verify the existence and correctness of G-ratio calculations
        expected_g_ratios_ctl1 = [0.78, 0.72, 0.97]  # G-ratios for CTL1
        expected_g_ratios_ctl2 = [0.38, 0.75, 0.76]  # G-ratios for CTL2
        pd.testing.assert_series_equal(result_df['G-Ratio_CTL1'], pd.Series(expected_g_ratios_ctl1, name='G-Ratio_CTL1'))
        pd.testing.assert_series_equal(result_df['G-Ratio_CTL2'], pd.Series(expected_g_ratios_ctl2, name='G-Ratio_CTL2'))
    def test_ensure_columns_exist(self):
        # Ensure all required columns are present before calling the function
        for ax_col, my_col in self.experiments:
            self.assertIn(ax_col, self.df.columns)
            self.assertIn(my_col, self.df.columns)

        # Optionally, run the calculation function after confirming column presence
        result_df = calculate_g_ratio(self.df.copy(), self.experiments)
        self.assertIn('G-Ratio_CTL1', result_df.columns)
        self.assertIn('G-Ratio_CTL2', result_df.columns)

if __name__ == '__main__':
    unittest.main()
