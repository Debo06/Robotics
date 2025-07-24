
import unittest
import pandas as pd
from datetime import datetime, timedelta
from email_inventory_forecast_automation import generate_forecast

class TestForecastAutomation(unittest.TestCase):

    def setUp(self):
        # Create sample sales data
        dates = pd.date_range(end=datetime.today(), periods=30)
        self.sales_df = pd.DataFrame({
            'Date': [d.strftime('%Y-%m-%d') for d in dates] * 2,
            'Product': ['Widget A'] * 30 + ['Widget B'] * 30,
            'Units Sold': [50] * 30 + [30] * 30
        })

        # Create sample inventory data
        self.inventory_df = pd.DataFrame({
            'Date': [(datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')] * 2,
            'Location': ['New York', 'Chicago'],
            'Product': ['Widget A', 'Widget B'],
            'Inventory Quantity': [100, 80]
        })

    def test_generate_forecast(self):
        forecast_df = generate_forecast(self.inventory_df, self.sales_df)
        self.assertFalse(forecast_df.empty)
        self.assertIn('Forecasted Units', forecast_df.columns)
        self.assertEqual(len(forecast_df), 180)  # 90 days x 2 products

        # Check that forecasts are integers
        self.assertTrue(all(forecast_df['Forecasted Units'].apply(lambda x: isinstance(x, (int, float)))))

if __name__ == '__main__':
    unittest.main()
