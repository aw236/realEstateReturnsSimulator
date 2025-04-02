import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# CONFIGURABLE VARIABLES
# Initial Investment Parameters
INITIAL_MARKET_VALUE = 300000
DOWN_PAYMENT = 60000
INITIAL_MONTHLY_HOA = 200
INITIAL_RENTAL_INCOME = 2000
MONTHS_TO_CALCULATE = 36

# Scenario Growth Rates
PESSIMISTIC_RENT_INCREASE = 0.01  # 1%
PESSIMISTIC_VALUE_INCREASE = 0.01  # 1%
PESSIMISTIC_HOA_INCREASE = 0.01  # 1%
PROBABLE_RENT_INCREASE = 0.02  # 2%
PROBABLE_VALUE_INCREASE = 0.02  # 2%
PROBABLE_HOA_INCREASE = 0.01  # 1%

# One-time Repairs (month: cost)
REPAIRS = {
    3: 1000,  # $1000 repair in month 3
    15: 500  # $500 repair in month 15
}

# Rental Income Changes (month: amount)
RENTAL_CHANGES = {
    13: 2200  # Rent increase to $2200 in month 13
}

# Google Sheets Configuration
# CREDENTIALS_FILE = 'your-credentials.json'
# USER_EMAIL = 'your-email@example.com'


# END OF CONFIGURABLE VARIABLES

class RealEstateInvestment:
    def __init__(self, market_value, down_payment, monthly_hoa,
                 rental_income, months_to_calculate):
        self.market_value = market_value
        self.down_payment = down_payment
        self.monthly_hoa = monthly_hoa
        self.rental_income = rental_income
        self.months = months_to_calculate
        self.repairs = {}

    def add_repair(self, month, cost):
        """Add a one-time repair cost for a specific month"""
        self.repairs[month] = cost

    def update_rental_income(self, month, amount):
        """Update rental income for a specific month"""
        self.rental_income[month] = amount

    def calculate_returns(self, scenario="probable"):
        """Calculate monthly cash-on-cash returns"""
        if scenario == "pessimistic":
            rent_increase = PESSIMISTIC_RENT_INCREASE
            value_increase = PESSIMISTIC_VALUE_INCREASE
            hoa_increase = PESSIMISTIC_HOA_INCREASE
        else:  # probable
            rent_increase = PROBABLE_RENT_INCREASE
            value_increase = PROBABLE_VALUE_INCREASE
            hoa_increase = PROBABLE_HOA_INCREASE

        results = []
        current_value = self.market_value
        current_hoa = self.monthly_hoa
        current_rent = self.rental_income.get(1, 0)

        for month in range(1, self.months + 1):
            # Apply annual increases at the start of each year
            if month % 12 == 1 and month > 1:
                current_value *= (1 + value_increase)
                current_hoa *= (1 + hoa_increase)
                current_rent *= (1 + rent_increase)

            # Check for specific rental income overrides
            if month in self.rental_income:
                current_rent = self.rental_income[month]

            # Calculate monthly cash flow
            repair_cost = self.repairs.get(month, 0)
            monthly_cash_flow = current_rent - current_hoa - repair_cost

            # Calculate cash-on-cash return (annualized)
            if self.down_payment > 0:
                coc_return = (monthly_cash_flow * 12) / self.down_payment * 100
            else:
                coc_return = 0

            results.append({
                'Month': month,
                'Property Value': round(current_value, 2),
                'Rental Income': round(current_rent, 2),
                'HOA Cost': round(current_hoa, 2),
                'Repair Cost': repair_cost,
                'Monthly Cash Flow': round(monthly_cash_flow, 2),
                'CoC Return (%)': round(coc_return, 2)
            })

        return pd.DataFrame(results)


def export_to_google_sheets(pessimistic_df, probable_df):
    """Export results to Google Sheets"""
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    spreadsheet = client.create(f'Property_Investment_Analysis_{timestamp}')
    spreadsheet.share(USER_EMAIL, perm_type='user', role='writer')

    pessimistic_sheet = spreadsheet.add_worksheet("Pessimistic Scenario", rows=100, cols=20)
    probable_sheet = spreadsheet.add_worksheet("Probable Scenario", rows=100, cols=20)
    spreadsheet.del_worksheet(spreadsheet.sheet1)

    pessimistic_sheet.update([pessimistic_df.columns.values.tolist()] +
                             pessimistic_df.values.tolist())
    probable_sheet.update([probable_df.columns.values.tolist()] +
                          probable_df.values.tolist())

    return spreadsheet.url


def main():
    investment = RealEstateInvestment(
        market_value=INITIAL_MARKET_VALUE,
        down_payment=DOWN_PAYMENT,
        monthly_hoa=INITIAL_MONTHLY_HOA,
        rental_income={1: INITIAL_RENTAL_INCOME},
        months_to_calculate=MONTHS_TO_CALCULATE
    )

    # Apply repairs
    for month, cost in REPAIRS.items():
        investment.add_repair(month, cost)

    # Apply rental changes
    for month, amount in RENTAL_CHANGES.items():
        investment.update_rental_income(month, amount)

    # Calculate returns
    pessimistic_results = investment.calculate_returns("pessimistic")
    probable_results = investment.calculate_returns("probable")

    # Print results
    print("Pessimistic Scenario:")
    print(pessimistic_results)
    print("\nProbable Scenario:")
    print(probable_results)

    # Export to Google Sheets
    try:
        sheet_url = export_to_google_sheets(pessimistic_results, probable_results)
        print(f"\nResults exported to Google Sheets: {sheet_url}")
    except Exception as e:
        print(f"Error exporting to Google Sheets: {e}")
        print("Make sure you have the correct credentials file and permissions")


if __name__ == "__main__":
    main()
