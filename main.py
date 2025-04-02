import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import re
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import urllib.parse

# CONFIGURABLE VARIABLES
INITIAL_MARKET_VALUE = 300000
DOWN_PAYMENT = 60000
INITIAL_MONTHLY_HOA = 200
INITIAL_RENTAL_INCOME = 2000
MONTHS_TO_CALCULATE = 36
LOAN_INTEREST_RATE = 0.05
LOAN_TERM_YEARS = 30
MONTHLY_BROKERAGE_FEE = 50
ANNUAL_PROPERTY_TAX = 3600

PESSIMISTIC_RENT_INCREASE = 0.01
PESSIMISTIC_VALUE_INCREASE = 0.01
PESSIMISTIC_HOA_INCREASE = 0.01
PROBABLE_RENT_INCREASE = 0.02
PROBABLE_VALUE_INCREASE = 0.02
PROBABLE_HOA_INCREASE = 0.01

REPAIRS = {3: 1000, 15: 500}
RENTAL_CHANGES = {13: 2200}

# CREDENTIALS_FILE = 'your-credentials.json'
# USER_EMAIL = 'your-email@example.com'
SHEET_NAME = 'Property_Investment_Analysis'


# END OF CONFIGURABLE VARIABLES

class RealEstateInvestment:
    def __init__(self, market_value, down_payment, monthly_hoa, rental_income,
                 months_to_calculate, interest_rate, loan_term_years,
                 monthly_brokerage_fee, annual_property_tax):
        self.market_value = market_value
        self.down_payment = down_payment
        self.monthly_hoa = monthly_hoa
        self.rental_income = rental_income
        self.months = months_to_calculate
        self.interest_rate = interest_rate
        self.loan_term_months = loan_term_years * 12
        self.monthly_brokerage_fee = monthly_brokerage_fee
        self.monthly_property_tax = annual_property_tax / 12
        self.repairs = {}
        self.loan_amount = market_value - down_payment
        self.monthly_payment = self.calculate_monthly_payment()

    def calculate_monthly_payment(self):
        monthly_rate = self.interest_rate / 12
        n = self.loan_term_months
        if monthly_rate == 0:
            return self.loan_amount / n
        return self.loan_amount * (monthly_rate * (1 + monthly_rate) ** n) / ((1 + monthly_rate) ** n - 1)

    def calculate_mortgage_components(self, month, remaining_balance):
        monthly_rate = self.interest_rate / 12
        interest = remaining_balance * monthly_rate
        principal = self.monthly_payment - interest
        return principal, interest

    def add_repair(self, month, cost):
        self.repairs[month] = cost

    def update_rental_income(self, month, amount):
        self.rental_income[month] = amount

    def calculate_returns(self, scenario="probable"):
        if scenario == "pessimistic":
            rent_increase = PESSIMISTIC_RENT_INCREASE
            value_increase = PESSIMISTIC_VALUE_INCREASE
            hoa_increase = PESSIMISTIC_HOA_INCREASE
        else:
            rent_increase = PROBABLE_RENT_INCREASE
            value_increase = PROBABLE_VALUE_INCREASE
            hoa_increase = PROBABLE_HOA_INCREASE

        results = []
        current_value = self.market_value
        current_hoa = self.monthly_hoa
        current_rent = self.rental_income.get(1, 0)
        remaining_balance = self.loan_amount

        for month in range(1, self.months + 1):
            if month % 12 == 1 and month > 1:
                current_value *= (1 + value_increase)
                current_hoa *= (1 + hoa_increase)
                current_rent *= (1 + rent_increase)

            if month in self.rental_income:
                current_rent = self.rental_income[month]

            repair_cost = self.repairs.get(month, 0)
            principal, interest = self.calculate_mortgage_components(month, remaining_balance)
            total_mortgage_payment = principal + interest + self.monthly_brokerage_fee

            monthly_cash_flow = (current_rent - current_hoa - repair_cost -
                                 total_mortgage_payment - self.monthly_property_tax)

            coc_return = (monthly_cash_flow * 12) / self.down_payment * 100 if self.down_payment > 0 else 0
            equity = current_value - remaining_balance

            results.append({
                'Month': month,
                'Property Value': round(current_value, 2),
                'Rental Income': round(current_rent, 2),
                'HOA Cost': round(current_hoa, 2),
                'Repair Cost': repair_cost,
                'Monthly Mortgage Payment': round(total_mortgage_payment, 2),
                'Principal': round(principal, 2),
                'Interest': round(interest, 2),
                'Brokerage': self.monthly_brokerage_fee,
                'Property Tax': round(self.monthly_property_tax, 2),
                'Monthly Cash Flow': round(monthly_cash_flow, 2),
                'CoC Return (%)': round(coc_return, 2),
                'Down Payment': self.down_payment,
                'Equity': round(equity, 2)
            })

            remaining_balance -= principal

        return pd.DataFrame(results)


def validate_email(email):
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, email))


def delete_existing_sheet(service, sheet_name):
    try:
        # Minimal query, let it fail gracefully if needed
        query = urllib.parse.quote(sheet_name)
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name, mimeType)',
            pageSize=10
        ).execute()
        files = results.get('files', [])
        for file in files:
            if file['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                print(f"Found existing sheet '{file['name']}' (ID: {file['id']}). Deleting...")
                service.files().delete(fileId=file['id']).execute()
                print(f"Deleted sheet '{file['name']}'.")
    except HttpError as e:
        print(f"Error checking/deleting existing sheet: {str(e)}")
        print("Continuing without deleting existing sheet...")


def export_to_google_sheets(pessimistic_df, probable_df):
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"Error: Credentials file '{CREDENTIALS_FILE}' not found.")
        return None

    if not validate_email(USER_EMAIL):
        print(f"Error: Invalid email address '{USER_EMAIL}'.")
        return None

    try:
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        drive_service = build('drive', 'v3', credentials=creds)

        # Attempt to delete, but proceed if it fails
        delete_existing_sheet(drive_service, SHEET_NAME)

        # Create new spreadsheet
        spreadsheet = client.create(SHEET_NAME)
        spreadsheet.share(USER_EMAIL, perm_type='user', role='writer')

        for df, sheet_name in [(pessimistic_df, "Pessimistic Scenario"),
                               (probable_df, "Probable Scenario")]:
            worksheet = spreadsheet.add_worksheet(sheet_name, rows=100, cols=20)

            # Convert DataFrame to list of lists, ensuring all values are strings
            data = [df.columns.tolist()] + df.values.tolist()
            data = [[str(cell) for cell in row] for row in data]

            # Update using update_cells instead
            cells = worksheet.range(1, 1, len(data), len(data[0]))
            for i, cell in enumerate(cells):
                row = i // len(data[0])
                col = i % len(data[0])
                cell.value = data[row][col]
            worksheet.update_cells(cells)

            # Apply currency format
            currency_columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'M']
            for col in currency_columns:
                worksheet.format(f"{col}2:{col}{len(df) + 1}", {
                    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0"}
                })

            # Adjust column widths
            for i, col in enumerate(df.columns):
                max_width = max(len(str(col)), max(len(str(x)) for x in df[col]))
                worksheet.spreadsheet.batch_update({
                    "requests": [{
                        "updateDimensionProperties": {
                            "range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": i,
                                      "endIndex": i + 1},
                            "properties": {"pixelSize": max_width * 8},
                            "fields": "pixelSize"
                        }
                    }]
                })

        spreadsheet.del_worksheet(spreadsheet.sheet1)
        return spreadsheet.url
    except Exception as e:
        print(f"Error exporting to Google Sheets: {str(e)}")
        return None


def main():
    investment = RealEstateInvestment(
        market_value=INITIAL_MARKET_VALUE,
        down_payment=DOWN_PAYMENT,
        monthly_hoa=INITIAL_MONTHLY_HOA,
        rental_income={1: INITIAL_RENTAL_INCOME},
        months_to_calculate=MONTHS_TO_CALCULATE,
        interest_rate=LOAN_INTEREST_RATE,
        loan_term_years=LOAN_TERM_YEARS,
        monthly_brokerage_fee=MONTHLY_BROKERAGE_FEE,
        annual_property_tax=ANNUAL_PROPERTY_TAX
    )

    for month, cost in REPAIRS.items():
        investment.add_repair(month, cost)

    for month, amount in RENTAL_CHANGES.items():
        investment.update_rental_income(month, amount)

    pessimistic_results = investment.calculate_returns("pessimistic")
    probable_results = investment.calculate_returns("probable")

    print("Pessimistic Scenario:")
    print(pessimistic_results)
    print("\nProbable Scenario:")
    print(probable_results)

    sheet_url = export_to_google_sheets(pessimistic_results, probable_results)
    if sheet_url:
        print(f"\nResults exported to Google Sheets: {sheet_url}")


if __name__ == "__main__":
    main()
