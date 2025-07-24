
import os
import pandas as pd
from datetime import datetime, timedelta
from win32com.client import Dispatch

def download_attachment_from_outlook(subject_keyword, download_folder):
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    for message in messages:
        if subject_keyword in message.Subject and message.Attachments.Count > 0:
            attachment = message.Attachments.Item(1)
            attachment.SaveAsFile(os.path.join(download_folder, attachment.FileName))
            print(f"Downloaded: {attachment.FileName}")
            return os.path.join(download_folder, attachment.FileName)
    raise FileNotFoundError("No matching email with attachment found.")

def load_data(inventory_file, sales_file):
    inventory_df = pd.read_excel(inventory_file)
    sales_df = pd.read_excel(sales_file)
    return inventory_df, sales_df

def generate_forecast(inventory_df, sales_df):
    forecast_df = sales_df.copy()
    forecast_df['Date'] = pd.to_datetime(forecast_df['Date'])
    last_date = forecast_df['Date'].max()
    forecast = []

    for product in forecast_df['Product'].unique():
        recent_sales = forecast_df[forecast_df['Product'] == product].tail(30)
        avg_daily_sales = recent_sales['Units Sold'].mean()
        for i in range(90):
            forecast.append({
                'Date': (last_date + timedelta(days=i+1)).strftime('%Y-%m-%d'),
                'Product': product,
                'Forecasted Units': round(avg_daily_sales)
            })

    forecast_df = pd.DataFrame(forecast)
    combined = pd.merge(inventory_df, forecast_df, on='Product', how='left')
    return combined

def save_forecast(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"Forecast saved to: {output_path}")

def main():
    subject = "Daily Inventory"
    download_path = os.getcwd()
    sales_file = "sales_data.xlsx"

    try:
        inventory_file = download_attachment_from_outlook(subject, download_path)
        inventory_df, sales_df = load_data(inventory_file, sales_file)
        forecast_df = generate_forecast(inventory_df, sales_df)
        save_forecast(forecast_df, "forecast_output.xlsx")
    except Exception as e:
        print(f"Automation failed: {e}")

if __name__ == "__main__":
    main()
