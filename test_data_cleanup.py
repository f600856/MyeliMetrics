import unittest
import pandas as pd
import numpy as np
from data_cleanup import data_cleanup 



class TestDataCleanup(unittest.TestCase):
    def test_data_cleanup(self):
        data = {
            'CTL1_Ax': [0.79, 0.85, 0.66],
            'CTL1_My': [0.22, 0.33, 0.02],
            'CTL2_Ax': [0.14, 0.54, 0.63],
            'CTL2_My': [0.23, 0.18, 0.20],
        }
        df = pd.DataFrame(data)
        experiments = [
            ('CTL1_Ax', 'CTL1_My'), 
            ('CTL2_Ax', 'CTL2_My')
        ]
        expected_data = {
            'CTL1_Ax': [0.79, 0.85, np.nan],
            'CTL1_My': [0.22, 0.33, np.nan],
            'CTL2_Ax': [0.54, 0.63, np.nan],
            'CTL2_My': [0.18, 0.20, np.nan]
        }
        expected_df = pd.DataFrame(expected_data)
        status, message = data_cleanup(df, experiments)
        self.assertEqual(status, "Success")
        self.assertEqual(message, "Data cleaned successfully.")
        pd.testing.assert_frame_equal(df, expected_df)



def run_tests():
    suite = unittest.TestLoader().loadTestsFromTestCase(TestDataCleanup)
    runner = unittest.TextTestRunner()
    result = runner.run(suite)
    if result.wasSuccessful():
        print("All tests passed successfully!")
    else:
        print("Some tests failed.")

if __name__ == '__main__':
    run_tests()
