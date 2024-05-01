import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import numpy_financial as npf

rent_roll_file = pd.read_excel('/Users/andriydavydyuk/Desktop/Underwriting_Project/Rent+Roll+Detail.xls', sheet_name=None, skiprows=5)
t12_file = pd.read_excel('/Users/andriydavydyuk/Desktop/Underwriting_Project/T12.xlsx', sheet_name=None, skiprows=6)
data_for_graphs = pd.read_csv('/Users/andriydavydyuk/Desktop/Underwriting_Project/data_for_graphs.csv', delimiter=',')
submarket_zipcode = pd.read_csv('/Users/andriydavydyuk/Desktop/Underwriting_Project/Zip_to_Submarket.csv', delimiter=',')
multifamily_file = pd.read_csv('/Users/andriydavydyuk/Desktop/Underwriting_Project/multifamily.csv', delimiter=',')

#########################__________Current In-Place Rents Table____________################################

# Access the DataFrame from the dictionary using the sheet name
rent_roll_sheet_name = 'Sheet1' 
rent_roll_df = rent_roll_file[rent_roll_sheet_name]

# Get unique values from Floorplan column to get units
units = rent_roll_df['Floorplan'].unique()
total_unit = rent_roll_df['Floorplan'].value_counts()
#total_unit_dictionary = rent_roll_df['Floorplan'].value_counts().to_dict()


# Filter out the rows with empty floorplan types
rent_roll_df = rent_roll_df[rent_roll_df['Floorplan'] != '']

# Initialize lists to store data
floorplans = []
number_of_units = []
square_feet = []
in_place_rents = []
market_values = []

# Iterate over each unique unit and its count
for unit, count in total_unit.items():
    # Filter rows where the unit matches
    filtered_rows = rent_roll_df[rent_roll_df['Floorplan'] == unit]
    
    # Access the corresponding square feet value from column F
    square_feet_value = filtered_rows['SQFT'].iloc[0] if not filtered_rows.empty else np.nan
    
    # Get the average lease plan for the current unit type 
    # Filter out zero values from 'Lease Rent' column before calculating mean
    in_place = rent_roll_df.loc[(rent_roll_df['Floorplan'] == unit) & (rent_roll_df['Lease Rent'] > 0), 'Lease Rent'].mean()
    in_place = 0 if pd.isnull(in_place) else in_place

    market = rent_roll_df.loc[(rent_roll_df['Floorplan'] == unit) & (rent_roll_df['Market + Addl.'] > 0), 'Market + Addl.'].mean()
    market = 0 if pd.isnull(market) else market
    
    # Append data to lists
    floorplans.append(unit)
    number_of_units.append(count)
    square_feet.append(square_feet_value)
    in_place_rents.append(in_place) 
    market_values.append(market)

# Convert lists to numpy arrays
market_values_array = np.array(market_values)
square_feet_array = np.array(square_feet)
in_place_rents_array = np.array(in_place_rents)

# Calculate market column based on the formula market_values/square_feet
price_per_SF = np.where(np.isnan(in_place_rents_array) | np.isnan(square_feet_array), 0, in_place_rents_array / square_feet_array)

# Create DataFrame from lists
in_place_rents_table_df = pd.DataFrame({
    '# of Units': number_of_units,
    'Type': floorplans,
    'SF': square_feet,
    'Rent/Month': in_place_rents,
    '$/SF': price_per_SF
})

# Sort DataFrame by 'Type' column
in_place_rents_table_df = in_place_rents_table_df.sort_values(by=['Type'], ascending=False)

# Define a function to classify the floorplan type
def classify_floorplan(floorplan):
    if 'E' in floorplan:
        return 'studio'
    elif 'A' in floorplan:
        return '1 bedroom'
    elif 'B' in floorplan:
        return '2 bedroom'
    else:
        return 'Unknown'

# classify the floorplans based on their types
classified_floorplans = [classify_floorplan(unit) for unit in floorplans]

# Add the classification column to the DataFrame
in_place_rents_table_df['Classification'] = classified_floorplans

# Calculate total values
total_units = in_place_rents_table_df['# of Units'].sum()
total_square_feet = in_place_rents_table_df['SF'].sum()
total_square_feet_all_units = np.sum(in_place_rents_table_df['# of Units'] * in_place_rents_table_df['SF']) / total_units if total_units != 0 else 0
total_rent_month_all_units = np.sum(in_place_rents_table_df['# of Units'] * in_place_rents_table_df['Rent/Month']) / total_units if total_units != 0 else 0
total_rent_month_per_sf = total_rent_month_all_units / total_square_feet_all_units if total_square_feet != 0 else 0

# Create a dictionary with total values
total_values = {
    '# of Units': total_units,
    'Type': '',
    'SF': total_square_feet_all_units,
    'Rent/Month': total_rent_month_all_units,
    '$/SF': total_rent_month_per_sf 
}

# Append the total values as the last row in the DataFrame
in_place_rents_table_df.loc[len(in_place_rents_table_df)] = total_values

# default stabelized rents table
stabilized_rent_df = pd.DataFrame({
    '# of Units': number_of_units,
    'Type': floorplans,
    'SF': square_feet,
    'Stabilized Rent/Month': market_values,
    'Stabilized $/SF': price_per_SF
})


#########################__________Stabilized Rents Table And Comps____________################################

zipcode = 76118
property_class = 'Class C'
quarter = '2023 Q3'
total_cost_per_unit = 432000 # change that to take from snapshot!

# Filter Data Frame by a given zipcode
submarkets_for_given_zipcode = submarket_zipcode.loc[submarket_zipcode['PostalCode'] == zipcode, 'SubmarketName']

# Condition to find vacancy rates by zipcode using the file of submarket to zipcode
if not submarkets_for_given_zipcode.empty:
    # Filter the DataFrame 'data_for_graphs' for submarkets associated with the given zipcode
    submarket_data_for_given_zipcode = data_for_graphs[(data_for_graphs['Submarket_Name'].isin(submarkets_for_given_zipcode)) & 
                                                       (data_for_graphs['Slice'] == property_class) &
                                                       (data_for_graphs['Period'] == quarter)]

    # Filter the DataFrame 'multifamily' for submarkets associated with the given zipcode
    multifamily_filtered = multifamily_file[(multifamily_file['Submarket'].isin(submarkets_for_given_zipcode)) & 
                                       (multifamily_file['Slice'] == property_class) & 
                                       (multifamily_file['Period'] == quarter)]
    
    if not submarket_data_for_given_zipcode.empty and not multifamily_filtered.empty:
        # Filter out vacancy rates that are 0 or null
        filtered_vacancy_rates = submarket_data_for_given_zipcode['Vacancy_Rate'].dropna().replace(0, np.nan)
        filtered_rent_unit = submarket_data_for_given_zipcode['Asking_Rent_Unit'].dropna().replace(0, np.nan)
        filtered_rent_sf = submarket_data_for_given_zipcode['Asking_Rent_SF'].dropna().replace(0, np.nan)
        
        # Extract cap rate from multifamily data
        cap_rate = multifamily_filtered['Market_Cap_Rate'].iloc[0]

        if not filtered_vacancy_rates.empty:
            # Calculate the average vacancy rate for submarkets associated with the given zipcode
            average_vacancy_rate = filtered_vacancy_rates.mean() * 100
            average_rent_unit = filtered_rent_unit.mean()
            average_rent_sf = filtered_rent_sf.mean()
            average_market_price = market_values_array.mean()
            cap_rate = cap_rate * 100

            print(f"Average vacancy rate for zipcode {zipcode}, {property_class}, {quarter}: {average_vacancy_rate:.2f}%")
            print(f"Average rent per unit for zipcode {zipcode}, {property_class}, {quarter}: {average_rent_unit:.2f}")
            print(f"Average rent per SF for zipcode {zipcode}, {property_class}, {quarter}: {average_rent_sf:.2f}")
            print(f"Cap rate for zipcode {zipcode}, {property_class}, {quarter}: {cap_rate:.2f}%")

            comps_table_df = pd.DataFrame({
                'Comps Metric': ['Average Vacancy Rate', 'Average Rent per Unit', 'Average Rent per SF', 'Market Price Rent Roll'],
                'Comps Value': [f'{average_vacancy_rate:.2f}%', f'${average_rent_unit:.2f}', f'${average_rent_sf:.2f}', f'${average_market_price:.2f}' ]
            })
                
            def adjust_price_rent_month(row):
                if average_vacancy_rate > 10:
                    if row['Rent/Month'] == 0:
                        return 0
                    else:
                        return (row['Rent/Month'] + market) / 2
                else:
                    return row['Rent/Month']
            
            def adjust_price_per_sf(row):
                if average_vacancy_rate > 10:
                    if row['Stabilized Rent/Month'] == 0:  # Assuming 'Stabilized Rent/Month' is the new stabilized rent column
                        return 0  # If rent is 0, set $/SF to 0
                    else:
                        return row['Stabilized Rent/Month'] / row['SF']  # Calculate $/SF based on stabilized rent and SF
                else:
                    return row['$/SF']  # If vacancy rate is low, keep the original $/SF
                
            # Apply the functions across the DataFrame
            #stabilized_rent_df['Stabilized Rent/Month'] = in_place_rents_table_df.apply(adjust_price_rent_month, axis=1)
            stabilized_rent_df['Stabilized Rent/Month'] = market_values_array    
            stabilized_rent_df['Stabilized $/SF'] = stabilized_rent_df.apply(adjust_price_per_sf, axis=1)

            stabilized_rent_df = stabilized_rent_df.sort_values(by=['Type'], ascending=False)

            total_rent_month_all_units_stabilized = np.sum(stabilized_rent_df['# of Units'] * stabilized_rent_df['Stabilized Rent/Month']) / total_units if total_units != 0 else 0
            total_rent_month_per_sf_stabilized = total_rent_month_all_units_stabilized / total_square_feet_all_units if total_square_feet != 0 else 0
            
            # Create a dictionary with total values
            total_values_stabilized = {
                '# of Units': total_units,
                'Type': '',
                'SF': total_square_feet_all_units,
                'Stabilized Rent/Month': total_rent_month_all_units_stabilized,
                'Stabilized $/SF': total_rent_month_per_sf_stabilized 
            }

            stabilized_rent_df['Classification'] = classified_floorplans

            # Append the total values as the last row in the DataFrame
            stabilized_rent_df.loc[len(in_place_rents_table_df)] = total_values_stabilized


        else:
            print(f"No non-zero or non-null vacancy rate data found for submarkets associated with zipcode {zipcode}")
    else:
        print(f"No vacancy rate data found for submarkets associated with zipcode {zipcode}")
else:
    print(f"No submarkets found for zipcode {zipcode}")

#########################__________Rent Premium Over In-Place Table____________################################

stabilized_rent_df.set_index('Type', inplace=False)
in_place_rents_table_df.set_index('Type', inplace=False)

# Calculating the Rent Premium Over In-Place
rent_premium = stabilized_rent_df['Stabilized Rent/Month'] - in_place_rents_table_df['Rent/Month']

rent_premium_df = pd.DataFrame({
    'Rent Premium Over In-Place': rent_premium
})

rent_premium_df = rent_premium_df.reindex(stabilized_rent_df.index)

total_rent_premium_over_in_place = total_values_stabilized['Stabilized Rent/Month'] - total_values['Rent/Month']

rent_premium_df.loc[len(in_place_rents_table_df)] = total_rent_premium_over_in_place


#########################__________Rent Increase Timing____________################################

# Determine Month Increase Starts (always 1)
month_increase_starts = 1 

# Check if any type contains 'P'
if any('P' in t for t in units):
    month_increase_complete = 24
else:
    month_increase_complete = 12

rent_increase_timing_df = pd.DataFrame({
    'Rent Increase Timing': '',
    'Month Increase Starts': [month_increase_starts],  # Convert to single value within a list
    'Month Increase Complete': [month_increase_complete]  # Convert to single value within a list
})

#########################__________Current Other Income____________################################

t12_sheet_name = '12 Month Trend'

t12_df = t12_file[t12_sheet_name]

total_values = t12_df.columns[-2] #column of totals

total_value_current_other_income = t12_df.loc[34, total_values]

current_other_income_df = pd.DataFrame({
    'Current Other Income' : '',
    'Laundry Income':0,
    'Late Fees':0,
    'Pet Rent':0,
    'Garage Income':0,
    'Deposit Forfeitures':0,
    'Other Charges':0,
    'TOTAL': [total_value_current_other_income]
})

#########################__________Yearly Growth Percentage Table_____________######################

annual_growth_rates_data = {
    'Metric': ['Market Rent Growth', 'Other Income Growth', 'Physical Vacancy', 'Concessions, Bad Debt & Loss to Lease'],
    'Year 1': [0.03, 0.03, 0.05, 0.03],  # 3% growth, 5% vacancy, etc.
    'Year 2': [0.03, 0.03, 0.05, 0.03],
    'Year 3': [0.03, 0.03, 0.05, 0.03],
    'Year 4': [0.03, 0.03, 0.05, 0.03],
    'Year 5': [0.03, 0.03, 0.05, 0.03],
    'Year 6': [0.03, 0.03, 0.05, 0.03],
    'Year 7': [0.03, 0.03, 0.05, 0.03],
    'Year 8': [0.03, 0.03, 0.05, 0.03],
    'Year 9': [0.03, 0.03, 0.05, 0.03],
    'Year 10': [0.03, 0.03, 0.05, 0.03],
    'Year 11': [0.03, 0.03, 0.05, 0.03]
}

# Create the DataFrame
annual_growth_rates_df = pd.DataFrame(annual_growth_rates_data)
annual_growth_rates_df.set_index('Metric', inplace=True)

#########################__________Utilities____________################################

gas_electricity_fees = t12_df.loc[100, total_values]
water_sewer_fees = t12_df.loc[101, total_values]
trash_fees = t12_df.loc[102, total_values]
other_utility_fees = t12_df.loc[103, total_values]
total_utilities_fees = gas_electricity_fees + water_sewer_fees + trash_fees + other_utility_fees

utilities_df = pd.DataFrame({
    'Utilities': '',
    'Gas & Electric': [gas_electricity_fees],
    'Water & Sewer': [water_sewer_fees],
    'Trash': [trash_fees],
    'Other': [other_utility_fees],
    'TOTAL': [total_utilities_fees]
}) 

#########################__________T-1 Annulized Revenue____________################################

# Find the last available month (excluding 'Total')
last_month = t12_df.columns[-3] 

# Get the value in the last month on the 6th row
rental_income = t12_df.loc[5, last_month]

# Multiply the value by 12 to get the annual rental income
annual_rental_income = rental_income * 12

vacancy_loss = t12_df.loc[10, last_month]

annual_vacancy_loss = vacancy_loss * 12

loss_to_lease = t12_df.loc[6, last_month]

concession = t12_df.loc[11, last_month]

bad_debt = t12_df.loc[12, last_month]

concession_baddebt_losstolease = (loss_to_lease + concession + bad_debt) * 12

net_rental_income = concession_baddebt_losstolease + annual_vacancy_loss + annual_rental_income

effective_gross_income = net_rental_income + total_value_current_other_income

t1_annulized_revenue_df = pd.DataFrame({
    'T-1 Annulized Revenue': '',
    'Gross Potential Rent': [annual_rental_income],
    'Vacancy': [annual_vacancy_loss],
    'Concessions, Bad Debt & Loss to Lease': [concession_baddebt_losstolease],
    'Net Rental Income': [net_rental_income],
    'Other Income': [total_value_current_other_income],
    'Effective Gross Income': [effective_gross_income]
})

#########################__________T-12 Operating Expenses____________################################

administrative_fees = t12_df.loc[48, total_values]

general_fees = t12_df.loc[56, total_values] + t12_df.loc[57, total_values] + t12_df.loc[58, total_values]

general_and_administrative = administrative_fees + general_fees

marketing_fees = t12_df.loc[52, total_values]

management_fees = t12_df.loc[55, total_values]

total_payroll = t12_df.loc[73, total_values]

turnover_make_ready = t12_df.loc[77, total_values]

repairs_maintenance = t12_df.loc[90, total_values] + t12_df.loc[97, total_values]

property_taxes = t12_df.loc[107, total_values]

insurance_fees = t12_df.loc[108, total_values]

total_expenses = general_and_administrative + marketing_fees + management_fees + total_payroll + turnover_make_ready + repairs_maintenance + property_taxes + insurance_fees + total_utilities_fees

net_operating_income = effective_gross_income - total_expenses

if total_units != 0:
    management_fees_per_unit = management_fees / total_units
    personnel_per_unit = total_payroll / total_units
    general_and_administrative_per_unit = general_and_administrative / total_units
    marketing_fees_per_unit = marketing_fees / total_units
    repairs_maintenance_per_unit = repairs_maintenance / total_units
    turnover_make_ready_per_unit = turnover_make_ready / total_units
    total_utilities_fees_per_unit = total_utilities_fees / total_units
    property_taxes_per_unit = property_taxes / total_units
    insurance_fees_per_unit = insurance_fees / total_units
    total_expenses_per_unit = total_expenses / total_units

else:
    management_fees_per_unit = 0  
    personnel_per_unit = 0
    general_and_administrative_per_unit = 0
    marketing_fees_per_unit = 0
    repairs_maintenance_per_unit = 0
    turnover_make_ready_per_unit = 0
    total_utilities_fees_per_unit = 0
    property_taxes_per_unit = 0
    insurance_fees_per_unit = 0
    total_expenses_per_unit = 0


t12_operating_expenses_df = pd.DataFrame({
    'T-12 Operating Expenses': ['Total', 'Per Unit'],
    'Management Fees': [management_fees, management_fees_per_unit ],
    'Personnel': [total_payroll, personnel_per_unit],
    'General & Administrative': [general_and_administrative, general_and_administrative_per_unit],
    'Marketing': [marketing_fees, marketing_fees_per_unit],
    'Repairs & Maintenance': [repairs_maintenance, repairs_maintenance_per_unit],
    'Turnover': [turnover_make_ready, turnover_make_ready_per_unit],
    'Contract Services': [0, 0],
    'Utilities': [total_utilities_fees,total_utilities_fees_per_unit],
    'Utility Reimbursements': [0, 0],
    'Property Taxes': [property_taxes, property_taxes_per_unit],
    'Insurance': [insurance_fees, insurance_fees_per_unit],
    'Capital Reserves': [0, 0],
    'TOTAL EXPENSES': [total_expenses, total_expenses_per_unit]
})

net_operating_income_df = pd.DataFrame({
    'Net Operating Income': [net_operating_income]
})

#########################__________Data By Months Table____________################################


# Calculate the totals for each DataFrame, which appear to be in the last row
# Here we assume that the last row contains the total value for rent/month and number of units.
total_units_in_place = in_place_rents_table_df.iloc[-1]['# of Units']
total_units_stabilized = stabilized_rent_df.iloc[-1]['# of Units']

# Ensure the total units are the same for both in-place and stabilized rents before proceeding
if total_units_in_place != total_units_stabilized:
    raise ValueError("The total number of units does not match between in-place and stabilized rents")

# Assuming an annual increase of 3%
annual_increase_rate = 0.03

# Create DataFrame for monthly data with explicit data types
months = 132
monthly_increase_data_df = pd.DataFrame({
    'Month': np.arange(1, months + 1),
    'Current In-Place Rents': 0.0,
    'Stabilized Rents': 0.0,
    'Percentage of Stabilized Units': 0.0,
    'Gross Potential Rent': 0.0
})

# Initialize the first month's rent based on total rent from respective tables
current_in_place_rents_initial = total_rent_month_all_units * total_units
stabilized_rents_initial = total_rent_month_all_units_stabilized * total_units

monthly_increase_data_df.at[0, 'Current In-Place Rents'] = current_in_place_rents_initial
monthly_increase_data_df.at[0, 'Stabilized Rents'] = stabilized_rents_initial

# Update each row in the DataFrame according to the stabilization schedule and rent increase
for index in range(1, months):
    monthly_increase_data_df.at[index, 'Current In-Place Rents'] = monthly_increase_data_df.at[index - 1, 'Current In-Place Rents'] * (1 + annual_increase_rate) ** (1/12)
    monthly_increase_data_df.at[index, 'Stabilized Rents'] = monthly_increase_data_df.at[index - 1, 'Stabilized Rents'] * (1 + annual_increase_rate) ** (1/12)

    month = monthly_increase_data_df.at[index, 'Month']
    if month < month_increase_starts:
        percentage_stabilized = 0.0
    elif month > month_increase_complete:
        percentage_stabilized = 1.0
    else:
        percentage_stabilized = (month - month_increase_starts) / (month_increase_complete - month_increase_starts)

    current_rent_total = monthly_increase_data_df.at[index, 'Current In-Place Rents']
    stabilized_rent_total = monthly_increase_data_df.at[index, 'Stabilized Rents']

    gross_potential_rent = (1 - percentage_stabilized) * current_rent_total + percentage_stabilized * stabilized_rent_total

    monthly_increase_data_df.at[index, 'Percentage of Stabilized Units'] = percentage_stabilized * 100
    monthly_increase_data_df.at[index, 'Gross Potential Rent'] = gross_potential_rent

gross_potential_rent_1_year = monthly_increase_data_df.loc[:11, 'Gross Potential Rent'].sum()
print('Gross Potential Rent for the 1 year is ', gross_potential_rent_1_year)

#########################__________Operating Expense Assumptions Table____________################################

if effective_gross_income != 0:
    management_fee_ratio = management_fees / effective_gross_income
else:
    management_fee_ratio = 'n.a.' 

# The maximum value between management fee ratio and 2.5%
max_value_management_fees = max(management_fee_ratio, 0.025)

#max_value_management_fees = max_value_management_fees * 100
#management_fee_ratio = management_fee_ratio * 100

operating_expense_assumptions_df = pd.DataFrame({
    'Operating Expense Assumptions': ['Actual', 'Projected'],
    'Annual Expense Growth': ['n.a.', 0.02 ],
    'Property Management Fee': [management_fee_ratio, max_value_management_fees]
})


print(operating_expense_assumptions_df)

#########################__________Property Tax Growth Rates____________################################

# Add logic to make changes in the rates based on US states!
# Define property tax growth rates for each year in a more accessible format
growth_rates = {
    'Year': range(1, 11),
    'Growth Rate': [0.05] + [0.03] * 9  # First year 5%, then 3% for the subsequent years
}

property_tax_growth_df = pd.DataFrame(growth_rates)
property_tax_growth_df.set_index('Year', inplace=True)


#########################__________Property Tax Information Table____________################################

tax_bill = 127911 # change based on a dataset
current_assessed_value = 5652861 # change based on a dataset
percentage_of_value_assessed = 90 # change based on a dataset
reassessed_upon_aquisition = 'No'
reassessed_upon_sale = 'No'

property_tax_rate = tax_bill / current_assessed_value

property_tax_information_df = pd.DataFrame({
    'Property Tax Information': [''],  # Assuming you meant to put some text here
    'Property Tax Rate': [property_tax_rate],
    'Tax Bill': [tax_bill],
    'Current Assessed Value': [current_assessed_value],
    'Reassessed Upon Acquisition': [reassessed_upon_aquisition],
    'Reassessed Upon Sale': [reassessed_upon_sale],
    'Percentage of Value Assessed': [percentage_of_value_assessed / 100]  # Converting percentage to a decimal
}, index=[0]) 

#########################__________HOLD PERIOD PROPERTY TAXES____________################################

# DataFrame to store yearly property tax information
columns = ['Year', 'Property Tax Rate', 'Tax Bill', 'Assessed Value', 'Non-Ad Valoreum Taxes', 'Less Credits', 'Total Property Taxes']
index = range(1, 11)  # Years 1 through 10
hold_period_property_tax_df = pd.DataFrame(index=index, columns=columns)
hold_period_property_tax_df['Year'] = hold_period_property_tax_df.index


# Calculate the tax information for each year
for year in index:
    growth_rate = property_tax_growth_df.loc[year, 'Growth Rate']
    if year == 1:
        assessed_value = current_assessed_value
    else:
        assessed_value = hold_period_property_tax_df.loc[year - 1, 'Assessed Value'] * (1 + growth_rate)
    
    tax_bill = assessed_value * property_tax_rate
    hold_period_property_tax_df.loc[year, 'Assessed Value'] = assessed_value
    hold_period_property_tax_df.loc[year, 'Tax Bill'] = tax_bill
    hold_period_property_tax_df.loc[year, 'Property Tax Rate'] = property_tax_rate
    hold_period_property_tax_df.loc[year, 'Non-Ad Valoreum Taxes'] = 0 # change
    hold_period_property_tax_df.loc[year, 'Less Credits'] = 0 # change
    hold_period_property_tax_df.loc[year, 'Total Property Taxes'] = property_tax_rate * assessed_value + hold_period_property_tax_df.loc[year, 'Non-Ad Valoreum Taxes'] - hold_period_property_tax_df.loc[year, 'Less Credits']

for year in range(1, 12):  # Example: setting data for 11 years
    hold_period_property_tax_df.loc[year - 1, 'Assessed Value'] = assessed_value  # 'year - 1' if 'year' starts from 1
    hold_period_property_tax_df.loc[year - 1, 'Tax Bill'] = tax_bill
    hold_period_property_tax_df.loc[year - 1, 'Property Tax Rate'] = property_tax_rate
    hold_period_property_tax_df.loc[year - 1, 'Non-Ad Valoreum Taxes'] = 0 
    hold_period_property_tax_df.loc[year - 1, 'Less Credits'] = 0
    hold_period_property_tax_df.loc[year - 1, 'Total Property Taxes'] = property_tax_rate * assessed_value + hold_period_property_tax_df.loc[year - 1, 'Non-Ad Valoreum Taxes'] - hold_period_property_tax_df.loc[year - 1, 'Less Credits']



hold_period_property_tax_df.reset_index(drop=True, inplace=True)

print(hold_period_property_tax_df)

#########################__________Projected Operating Expenses Table____________################################

# default values:
projected_total_payroll = total_payroll
projected_general_and_administrative = general_and_administrative
projected_marketing_fees = marketing_fees 
projected_repairs_maintenance = repairs_maintenance


# Constants for payroll calculations
full_time_worker_cost = 50000
half_time_assistant_cost = 25000
units_per_full_time_worker = 75

# Calculate number of full-time workers and remainder units
full_time_workers = total_units // units_per_full_time_worker
remainder_units = total_units % units_per_full_time_worker

# Determine the payroll based on number of full-time workers
projected_total_payroll = full_time_workers * full_time_worker_cost

# If there are any remainder units that do not complete another full 75, add a half-time assistant
if remainder_units > 5:
    projected_total_payroll += half_time_assistant_cost

projected_general_and_administrative = 200 * total_units

year_1_property_taxes = hold_period_property_tax_df.loc[1, 'Total Property Taxes']

projected_property_taxes = year_1_property_taxes


# Determine fee per unit based on property class
if property_class == 'Class A':
    marketing_fee_per_unit = 200
    maintenance_fee_per_unit = 0.02
    projected_contact_services_per_unit = 300
    capital_reserves_per_unit = 350 # not sure about the rate
elif property_class == 'Class B':
    marketing_fee_per_unit = 150
    maintenance_fee_per_unit = 0.03
    projected_contact_services_per_unit = 350
    capital_reserves_per_unit = 300 # not sure about the rate
elif property_class == 'Class C':
    marketing_fee_per_unit = 100
    maintenance_fee_per_unit = 0.04
    projected_contact_services_per_unit = 400
    capital_reserves_per_unit = 250 # not sure about the rate
else: # Default case if none of A, B, or C
    marketing_fee_per_unit = 0 
    maintenance_fee_per_unit = 0 
    projected_contact_services_per_unit = 0
    capital_reserves_per_unit = 0

# Calculate projected fees
projected_marketing_fees = marketing_fee_per_unit * total_units
projected_repairs_maintenance = maintenance_fee_per_unit * gross_potential_rent_1_year # effective gross income from Projected Revenue!!! CHANGE 1 after Annual CF
projected_contact_services = projected_contact_services_per_unit * total_units
projected_capital_reserves = capital_reserves_per_unit * total_units
projected_total_expenses = management_fees + projected_total_payroll + projected_general_and_administrative + projected_marketing_fees + projected_repairs_maintenance + projected_contact_services + turnover_make_ready + total_utilities_fees + year_1_property_taxes + insurance_fees + projected_capital_reserves
projected_turnover = 1000 * 0.25 * total_units # ????? why 1000? maybe to take the initial value of turnover_make_ready

projected_total_expenses_per_unit = projected_total_expenses / total_units
projected_property_taxes_per_unit = projected_property_taxes / total_units

projected_operating_expenses_df = pd.DataFrame({
    'Projected Operating Expenses': ['Total', 'Per Unit'],
    'Management Fees': [management_fees, management_fees_per_unit ], # change based on 'Monthly CF'!
    'Personnel': [projected_total_payroll, personnel_per_unit],
    'General & Administrative': [projected_general_and_administrative, general_and_administrative_per_unit],
    'Marketing': [projected_marketing_fees, marketing_fees_per_unit],
    'Repairs & Maintenance': [projected_repairs_maintenance, repairs_maintenance_per_unit],
    'Contract Services': [projected_contact_services, projected_contact_services_per_unit], # do not change projected_contact_services_per_unit
    'Turnover': [projected_turnover, turnover_make_ready_per_unit],
    'Utilities': [total_utilities_fees,total_utilities_fees_per_unit], # do not change
    'Utility Reimbursements': [0, 0],
    'Property Taxes': [projected_property_taxes, projected_property_taxes_per_unit ], 
    'Insurance': [insurance_fees, insurance_fees_per_unit],
    'Capital Reserves': [projected_capital_reserves, capital_reserves_per_unit], # change based on other rates (probably they are different by class)
    'TOTAL EXPENSES': [projected_total_expenses, total_expenses_per_unit]
})

##################################!!!!!!!!!!!!!!!!!!!________________Snapshot Sheet____________!!!!!!!!!!!!!!!!###############################

net_rentable_square_feet = total_square_feet_all_units * total_units
purchase_price = 5000000 # input of brokers data!
LTV_acquisition_info = 70 / 100
operating_reserves_working_capital = 100 * total_units
upfront_capex = 1000 * total_units
loan_fee = 1 / 100
loan_amount = purchase_price * LTV_acquisition_info
acquisition_loan_fees = loan_fee * loan_amount
amortization_years = 30
interest_only_period_months = 12
interest_rates_acquisition_loan = [0.0625, 0.0]
asset_management_fee = 3 / 100
refinance_month = 12
interest_rate_refinance = 0.055
amortization_years_refinance = 35
hold_period = 60
LTV_refinance = 80 / 100
closing_costs_sale_price_percentage = 3 / 100
yearly_exit_cap_rate_increment = 5 / 100
exit_cap_rate = cap_rate + (yearly_exit_cap_rate_increment / 12 * hold_period)

try:
    total_capex_budget_per_unit = (asset_management_fee + upfront_capex) / total_units
except ZeroDivisionError:
    total_capex_budget_per_unit = 0

property_detail_df = pd.DataFrame({
    'Property Details': '',
    'Year Built': 0, # change it with the input from brokers
    'Number of Units': [total_units],
    'Net Rentable Square Feet': [net_rentable_square_feet],
    'Average Unit Size': [total_square_feet_all_units]
})

acquisition_information_df = pd.DataFrame({
    'Asking Price': [0],
    'Purchase Price': [purchase_price], 
    'Purchase Price Per Unit': [0],
    'Acquisition Fee': [0],
    'Closing Costs': [0],
    'Cost of Capital': [0],
    'Disposition Fee': [0],
    'Operating Reserves/Working Capital': [operating_reserves_working_capital],
    'Upfront Capex': [upfront_capex],
    'Going-In Cap Rate': [0] 
})

acquisition_information_df['Purchase Price Per Unit'] = purchase_price / total_units
acquisition_information_df['Acquisition Fee'] = acquisition_information_df['Purchase Price'] * 0.03
acquisition_information_df['Closing Costs'] = acquisition_information_df['Purchase Price'] * 0.03
acquisition_information_df['Cost of Capital'] = 0.035 * (purchase_price * (1 - LTV_acquisition_info) + acquisition_information_df['Acquisition Fee'] + acquisition_information_df['Closing Costs'] + operating_reserves_working_capital + upfront_capex + acquisition_loan_fees)

acquisition_loan_information_df = pd.DataFrame({
    'Loan Amount': [loan_amount],
    'LTV': [LTV_acquisition_info],
    'Going-In DSCR': [0],
    'Going-In Debt Yield': [0],
    'Interest Rate': [max(interest_rates_acquisition_loan)],
    'Interest Only Period': [interest_only_period_months],
    'Loan Fee': [loan_fee],
    'Amortization': [amortization_years]
})

annual_interest_rate = sum(interest_rates_acquisition_loan) / 12

# Check for division by zero scenario
if annual_interest_rate == 0:
    print("Interest rate cannot be zero for PMT calculation.")
else:
    # Calculate the payment using numpy's pmt function, use numpy financial library
    monthly_payment = npf.pmt(annual_interest_rate, amortization_years * 12, -loan_amount)

    # Annual payment
    annual_payment = monthly_payment * 12

    try:
        acquisition_loan_information_df['Going-In DSCR'] = (net_operating_income * 12) / annual_payment if annual_payment != 0 else 0
    except ZeroDivisionError:
        acquisition_loan_information_df['Going-In DSCR'] = 0

    print(f"The calculated Going-In DSCR is: {acquisition_loan_information_df['Going-In DSCR']}")

try:
    acquisition_loan_information_df['Going-In Debt Yield'] = (net_operating_income * 12) / loan_amount if loan_amount != 0 else 0
except ZeroDivisionError:
    acquisition_loan_information_df['Going-In Debt Yield'] = 0

print(f"The calculated Going-In Debt Yield is: {acquisition_loan_information_df['Going-In Debt Yield']}")


refinance_information_df = pd.DataFrame({
    'Refinance Information': '',
    'Refinance Month': [refinance_month], #change to 24 if the 'Gross Proceeds From Refinance' is negative
    'Interest Rate': [interest_rate_refinance],
    'Interest Only Period': '0 Months',
    'Amortization': [amortization_years_refinance],

    'Cap Rate Used for Valuation, %': [cap_rate],
    'LTV': LTV_refinance,
    'Min DSCR': [1.18],
    'Refi NOI': 0,
    'Max Allowed Loan Amount at 1.18x DSCR': [0],
    'Going-In DSCR': [0],
    'Going-In Debt Yield': [0],

    'Loan Amount': [0],
    'Refinance Loan Origination Fee': [0],
    'Initial Loan Prepayment Penalty': [0],
    'Total Refinance Cost': [0]
})

exit_assumptions_df = pd.DataFrame({
    'Exit Assumptions': '',
    'Closing Costs (as percentage of Sale Price)': [closing_costs_sale_price_percentage],
    'Yearly Exit Cap Rate Increment': [yearly_exit_cap_rate_increment],
    'Exit Cap Rate': [exit_cap_rate],
    'Hold Period': [hold_period]
})

#HOLD = exit_assumptions_df['Hold Period'].iloc[0]

capital_and_operating_details_df = pd.DataFrame({
    'Asset Management Fee': [asset_management_fee],
    'Total Capex Budget': [0], # change after Scheduled Capex!
    'Total Capex Budget Per Unit': [total_capex_budget_per_unit],
    'Return on Cost': [0]
})

try:
    # Calculate the denominator
    denominator = capital_and_operating_details_df['Total Capex Budget'] - 0
    
    # Check if the denominator is zero or very close to zero
    capital_and_operating_details_df['Return on Cost'] = np.where(
        denominator > 0.001,  # Threshold to avoid division by very small number
        total_rent_premium_over_in_place * 12 * total_units / denominator,
        0  # Set to 0 if the denominator is zero or too small
    )
except Exception as e:
    print(f"An error occurred: {e}")
    capital_and_operating_details_df['Return on Cost'] = 0

total_uses_at_close = purchase_price + acquisition_information_df['Acquisition Fee'] + acquisition_information_df['Closing Costs'] + acquisition_information_df['Cost of Capital'] + operating_reserves_working_capital + upfront_capex + acquisition_loan_fees

total_uses_at_close_df = pd.DataFrame({
    'Total Uses (at Close)': '',
    'Purchase Price': purchase_price,
    'Acquisition Fee': acquisition_information_df['Acquisition Fee'],
    'Closing Costs': acquisition_information_df['Closing Costs'],
    'Cost of Capital':  acquisition_information_df['Cost of Capital'],
    'Operating Reserves/Working Capital': [operating_reserves_working_capital],
    'Repairs': [upfront_capex],
    'Acquisition Loan Fees': [acquisition_loan_fees],
    'Total': total_uses_at_close
})

lp_equity = total_uses_at_close - acquisition_loan_fees

sources_at_close_df = pd.DataFrame({
    'Sources (at Close)': '',
    'Acquisition Loan Proceeds': loan_amount,
    'LP Equity': lp_equity,
    'GP Equity': 0,
    'Total': total_uses_at_close
})  

sale_details_df = pd.DataFrame({
    'Sale Details': '',
    'Sale Price': [0] 
})  



##################################!!!!!!!!!!!!!!!!!!!________________Monthly CF Sheet____________!!!!!!!!!!!!!!!!###############################

# Start and end dates for the table creation. 
# The number of months in monthly_cf and monthly_increase_data_df MUST MATCH!
start_date = datetime.strptime("01.01.24", "%d.%m.%y")
end_date = datetime.strptime("31.12.34", "%d.%m.%y")

# Generate the date range
months = pd.date_range(start_date, end_date, freq='M')

# Categories and default values
categories_monthly_cf = {
    'Date': months,
    'Revenue': np.full(len(months), "", dtype=str), 
    'Gross Potential Rent': np.zeros(len(months)),
    'Physical Vacancy': np.zeros(len(months)),
    'Concessions, Bad Debt, Loss to Lease': np.zeros(len(months)),
    'Net Rental Income': np.zeros(len(months)),
    'Other Income': np.full(len(months), "", dtype=str),
    'Electricity': np.zeros(len(months)),
    'Water/Sewer': np.zeros(len(months)),
    'Pet Rent': np.zeros(len(months)),
    'Garage Income': np.zeros(len(months)),
    'Deposit Forfeitures': np.zeros(len(months)),
    'Other Charges': np.zeros(len(months)),
    'Total Other Income': np.zeros(len(months)),
    'Effective Gross Revenue': np.zeros(len(months)),
    'Expenses': np.full(len(months), "", dtype=str),
    'Management Fee': np.zeros(len(months)),
    'Personnel': np.zeros(len(months)),
    'General & Administrative': np.zeros(len(months)),
    'Marketing': np.zeros(len(months)),
    'Repairs & Maintenance': np.zeros(len(months)),
    'Turnover': np.zeros(len(months)),
    'Contract Services': np.zeros(len(months)),
    'Utilities': np.zeros(len(months)),
    'Utility Reimbursements': np.zeros(len(months)),
    'Property Taxes': np.zeros(len(months)),
    'Insurance': np.zeros(len(months)),
    'Capital Reserves': np.zeros(len(months)),
    'Total Expenses': np.zeros(len(months)),
    'Net Operating Income': np.zeros(len(months)),
    'Capital & Partnership Expenses': np.full(len(months), "", dtype=str),
    'Construction Expenses': np.zeros(len(months)), # after Sceduled Capex sheet
    'Construction Management Fee': np.zeros(len(months)), # after Snapshot sheet
    'Asset Management Fee': np.zeros(len(months)), # after Snapshot sheet
    'Cash Flow Before Debt Service': np.zeros(len(months)), # after Snapshot sheet
    'Debt Service': np.full(len(months), "", dtype=str),
    'Interest Payment': np.zeros(len(months)), # after Debt sheet
    'Principal Payment': np.zeros(len(months)), # after Debt sheet
    'Total Debt Service': np.zeros(len(months)), # after Debt sheet
    'Cash Flow After Debt Service': np.zeros(len(months)), # after Debt sheet
    'Acquisition & Sale Information': np.full(len(months), "", dtype=str),
    'Purchase Price': np.zeros(len(months)),  # after Snapshot sheet (broker input)
    'Acquisition Fee': np.zeros(len(months)), # after Snapshot sheet (broker input)
    'Closing Costs': np.zeros(len(months)), # after Snapshot sheet
    'Upfront Capex': np.zeros(len(months)), # after Snapshot sheet
    'Sale Price': np.zeros(len(months)),
    'Disposition Fee': np.zeros(len(months)), # after Snapshot sheet
    'Costs of Sale': np.zeros(len(months)), # after Snapshot sheet
    'Unlevered Net Cash Flow': np.zeros(len(months)), # after Snapshot sheet
    'Loan Information': np.full(len(months), "", dtype=str),
    'Loan Funding': np.zeros(len(months)), # after Snapshot sheet
    'Loan Payoff': np.zeros(len(months)), # after Snapshot sheet
    'Loan Fees': np.zeros(len(months)), # after Snapshot sheet
    'Working Construction Capital': np.full(len(months), "", dtype=str), 
    'Working Capital Contributed': np.zeros(len(months)), # after Snapshot sheet
    'Working Capital Distributed': np.zeros(len(months)), # after Snapshot sheet
    'Levered Net Cash Flow': np.zeros(len(months)), # after Snapshot sheet
    'Ending Working Capital Balance': np.zeros(len(months)), # after Snapshot sheet
    'Net Sales/Refinance Proceeds': np.zeros(len(months)) # after Snapshot sheet
}

monthly_cf_df = pd.DataFrame(categories_monthly_cf)

monthly_cf_df['Date'] = monthly_cf_df['Date'].dt.strftime('%Y-%m')  
monthly_cf_df['Month'] = np.arange(0, len(months))  # Adding a Month Number that iterates from 1 to 132

# Verify if the 'Date' is datetime type, if not convert it
if not pd.api.types.is_datetime64_any_dtype(monthly_cf_df['Date']):
    monthly_cf_df['Date'] = pd.to_datetime(monthly_cf_df['Date'])

# This part of the code generates the year index correctly for use with your DataFrame:
monthly_cf_df['Year'] = monthly_cf_df['Date'].dt.year - start_date.year + 1
monthly_cf_df['Year'] = monthly_cf_df['Year'].apply(lambda x: f'Year {x}' if x <= 11 else 'Year 11')
monthly_cf_df['Year'] = monthly_cf_df['Year'].apply(lambda x: int(x.split()[1]) if isinstance(x, str) else x)

# Determine the base year (the year of the first entry)
base_year = monthly_cf_df['Year'].min()

# Functions and calculations for df values:

# Ensure your main DataFrame setup matches the range exactly
if len(monthly_cf_df) == len(monthly_increase_data_df):
    monthly_cf_df['Gross Potential Rent'] = monthly_increase_data_df['Gross Potential Rent']
    if monthly_cf_df['Gross Potential Rent'].equals(monthly_increase_data_df['Gross Potential Rent']):
        print("Update successful, data matched.")
    else:
        print("Data mismatch after assignment!")
else:
    print("Length mismatch: monthly_cf_df has", len(monthly_cf_df), "rows; monthly_increase_data_df has", len(monthly_increase_data_df), "rows")

def calculate_physical_vacancy(row):
    try:
        year_column = row['Year'] if row['Year'] in annual_growth_rates_df.columns else 'Year 11'
        vacancy_rate = annual_growth_rates_df.at['Physical Vacancy', year_column]
        result = -row['Gross Potential Rent'] * vacancy_rate
        return result
    except Exception as e:
        print(f"Error processing row with Year {row['Year']}: {e}")
        return np.nan  # Use np.nan as a fallback in case of error
    
def calculate_value(row):
    try:
        # Sum of 'Gross Potential Rent' and 'Physical Vacancy'
        total = row['Gross Potential Rent'] + row['Physical Vacancy']

        # Determine the year and ensure it's within range
        year_column = row['Year'] if row['Year'] in annual_growth_rates_df.columns else 'Year 11'
        
        # Get the multiplier from the fifth row of the defined range, here assumed to be 'Some Metric'
        multiplier = annual_growth_rates_df.at['Concessions, Bad Debt & Loss to Lease', year_column]
        
        # Calculate the final value
        result = -total * multiplier
        return result
    except Exception as e:
        print(f"Error processing row with Year {row['Year']}: {e}")
        return np.nan  # Use np.nan as a fallback in case of error
    
monthly_cf_df.loc[1, 'Other Charges'] = current_other_income_df.loc[0, 'TOTAL'] / 12

# Calculate the 'Total Other Income' for subsequent months
for month in range(2, len(monthly_cf_df)):  # This starts from the second month as intended
    previous_value = monthly_cf_df.loc[month - 1, 'Other Charges']
    year_label = monthly_cf_df.loc[month, 'Year']

    # Ensure the year is within the defined year bounds of annual_growth_rates_df
    if year_label not in annual_growth_rates_df.columns:
        year_label = 'Year 11'  

    growth_rate = annual_growth_rates_df.at['Other Income Growth', year_label]
    monthly_cf_df.loc[month, 'Other Charges'] = previous_value * (1 + growth_rate)**(1/12)

projected_management_fee = -operating_expense_assumptions_df.loc[operating_expense_assumptions_df['Operating Expense Assumptions'] == 'Projected', 'Property Management Fee'].iloc[0]

annual_growth_rate = operating_expense_assumptions_df.loc[
    operating_expense_assumptions_df['Operating Expense Assumptions'] == 'Projected', 'Annual Expense Growth'].iloc[0]

annual_personnel_cost = projected_operating_expenses_df['Personnel'].iloc[0]
monthly_personnel_base = -annual_personnel_cost / 12

annual_general_administative_cost = projected_operating_expenses_df['General & Administrative'].iloc[0]
monthly_general_administative_base = -annual_general_administative_cost / 12

annual_marketing_cost = projected_operating_expenses_df['Marketing'].iloc[0]
monthly_marketing_base = -annual_marketing_cost/ 12

annual_repairs_maintenance_cost = projected_operating_expenses_df['Repairs & Maintenance'].iloc[0]
monthly_repairs_maintenance_base = -annual_repairs_maintenance_cost/ 12

annual_turnover_cost = projected_operating_expenses_df['Turnover'].iloc[0]
monthly_turnover_base = -annual_turnover_cost/ 12

annual_contract_services_cost = projected_operating_expenses_df['Contract Services'].iloc[0]
monthly_contract_services_base = -annual_contract_services_cost/ 12

annual_utilities_cost = projected_operating_expenses_df['Utilities'].iloc[0]
monthly_utilities_base = -annual_utilities_cost/ 12

annual_utility_reimbursements_cost = projected_operating_expenses_df['Utility Reimbursements'].iloc[0]
monthly_utility_reimbursements_base = -annual_utility_reimbursements_cost/ 12

annual_insurance_cost = projected_operating_expenses_df['Insurance'].iloc[0]
monthly_insurance_base = -annual_insurance_cost/ 12

annual_capital_reserves_cost = projected_operating_expenses_df['Capital Reserves'].iloc[0]
monthly_capital_reserves_base = -annual_capital_reserves_cost/ 12


# Apply the functions to calculate the df values:

monthly_cf_df['Physical Vacancy'] = monthly_cf_df.apply(calculate_physical_vacancy, axis=1)
monthly_cf_df['Concessions, Bad Debt, Loss to Lease'] = monthly_cf_df.apply(calculate_value, axis=1)
monthly_cf_df['Net Rental Income'] = monthly_cf_df['Gross Potential Rent'] + monthly_cf_df['Physical Vacancy'] + monthly_cf_df['Concessions, Bad Debt, Loss to Lease']
monthly_cf_df['Total Other Income'] = monthly_cf_df['Electricity'] + monthly_cf_df['Water/Sewer'] + monthly_cf_df['Pet Rent'] + monthly_cf_df['Garage Income'] + monthly_cf_df['Deposit Forfeitures'] + monthly_cf_df['Other Charges']
monthly_cf_df['Effective Gross Revenue'] = monthly_cf_df['Total Other Income'] + monthly_cf_df['Net Rental Income']

monthly_cf_df['Management Fee'] = monthly_cf_df['Effective Gross Revenue'] * projected_management_fee
monthly_cf_df['Personnel'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_personnel_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)
monthly_cf_df['General & Administrative'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_general_administative_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)
monthly_cf_df['Marketing'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_marketing_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)
monthly_cf_df['Repairs & Maintenance'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_repairs_maintenance_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
) # CHANGE in 'projected_repairs_maintenance = maintenance_fee_per_unit * 1' after Annual CF!!!

monthly_cf_df['Turnover'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_turnover_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)

monthly_cf_df['Contract Services'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_contract_services_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)

monthly_cf_df['Utilities'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_utilities_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)

monthly_cf_df['Utility Reimbursements'] = monthly_cf_df['Month'].apply(
    lambda month: 0 if month == 0 else monthly_utility_reimbursements_base * ((1 + annual_growth_rate) ** ((month - 1) / 12))
)

years = monthly_cf_df['Year'].unique()
yearly_factors = {year: (1 + annual_growth_rate) ** (year) for year in years}

# Conditionally apply these factors, starting from Month 1
monthly_cf_df['Insurance'] = monthly_cf_df.apply(
    lambda row: monthly_insurance_base * yearly_factors[row['Year']] if row['Month'] > 0 else 0, axis=1
)

monthly_cf_df['Capital Reserves'] = monthly_cf_df.apply(
    lambda row: monthly_capital_reserves_base * yearly_factors[row['Year']] if row['Month'] > 0 else 0, axis=1
)


#monthly_cf_df['Net Operating Income'] = monthly_cf_df['Effective Gross Revenue'] + projected_total_expenses

monthly_cf_df['Property Taxes'] = monthly_cf_df.apply(
    lambda row: 0 if row['Month'] == 0 else -projected_operating_expenses_df.loc[0, 'Property Taxes'] / 12,
    axis=1
)

monthly_cf_df['Total Expenses'] =  monthly_cf_df['Management Fee'] + monthly_cf_df['Personnel'] +  monthly_cf_df['General & Administrative'] + monthly_cf_df['Marketing'] + monthly_cf_df['Repairs & Maintenance'] + monthly_cf_df['Turnover'] + monthly_cf_df['Contract Services'] + monthly_cf_df['Utilities'] + monthly_cf_df['Utility Reimbursements'] + monthly_cf_df['Insurance'] + monthly_cf_df['Property Taxes'] + monthly_cf_df['Capital Reserves']

monthly_cf_df['Net Operating Income'] = monthly_cf_df.apply(
    lambda row: 0 if row['Month'] == 0 else row['Effective Gross Revenue'] + row['Total Expenses'],
    axis=1
)

sale_price = []
for i in range(len(monthly_cf_df)):
    current_month = monthly_cf_df.iloc[i]['Month']

    if current_month == hold_period:
        print("Hold period matched at Month:", current_month)  # Confirm hold period match
        if i + 12 < len(monthly_cf_df):  # Ensure there are 12 months available after the current month
            noi_sum = monthly_cf_df.iloc[i + 1: i + 13]['Net Operating Income'].sum()
            print("Sum of NOI from", i + 1, "to", i + 12, ":", noi_sum)  # Display calculated sum
            calculated_value = noi_sum / (exit_cap_rate/100)
            print("Calculated Sale Price:", calculated_value)  # Show the resulting sale price
            sale_price.append(calculated_value)
        else:
            sale_price.append(0)
            print("Not enough data to calculate from month", i + 1)  # Indicate insufficient data
        sale_price.extend([0] * (len(monthly_cf_df) - i - 1))
        break
    else:
        sale_price.append(0)

sale_details_df['Sale Price'] = sale_details_df['Sale Price'].astype(float)
monthly_cf_df['Sale Price'] = sale_price

# Update the Sale Price in Snapshot
total_sale_price = monthly_cf_df['Sale Price'].sum()
total_sale_price = float(total_sale_price)
sale_details_df.at[0, 'Sale Price'] = total_sale_price
total_sale_price = int(round(total_sale_price))

# from the Acquisition Information Data Frame to get the sale price:
acquisition_information_df['Disposition Fee'] = 0.01 * total_sale_price

'''monthly_cf_df['Property Taxes'] = monthly_cf_df.apply(
    lambda row: (
        (total_sale_price * (percentage_of_value_assessed / 100) * property_tax_rate +
         hold_period_property_tax_df.loc[int(row['Year']) - 1, 'Total Property Taxes'] -
         hold_period_property_tax_df.loc[int(row['Year']) - 1, 'Less Credits']) / 12 * -1
        if (row['Month'] > hold_period and reassessed_upon_sale == "Yes")
        else hold_period_property_tax_df.loc[int(row['Year']) - 1, 'Total Property Taxes'] / 12 * -1
        if row['Month'] > 0 else 0
    ), axis=1
)'''


# from Snapshot, refinance information:
try:
    # We will use boolean indexing to filter the months
    start_month = refinance_month
    end_month = refinance_month + 12

    # Filtering the DataFrame for the specific range and summing the 'Net Operating Income'
    refinance_information_df['Refi NOI'] = monthly_cf_df.loc[start_month:end_month, 'Net Operating Income'].sum()
    print("Sum of Net Operating Income from month", start_month, "to", end_month, ":", refinance_information_df['Refi NOI'])

except KeyError as e:
    print("An error occurred:", e)

periodic_payment = -refinance_information_df['Refi NOI'].iloc[0] / refinance_information_df['Min DSCR'].iloc[0]
refinance_information_df['Max Allowed Loan Amount at 1.18x DSCR'] = npf.pv(rate=interest_rate_refinance, nper=amortization_years_refinance, pmt=periodic_payment)


# calculation of Going-In DSCR , refinance information:
monthly_interest_rate = sum(interest_rates_acquisition_loan) / 12
total_periods = amortization_years * 12
monthly_payment = npf.pmt(monthly_interest_rate, total_periods, -loan_amount)

try:
    # Index for the DataFrame should be adjusted to zero-based index, and check if it exists
    refinance_index = refinance_month  # Since refinance_month is supposed to be an index, adjust if necessary
    if 'Net Operating Income' in monthly_cf_df.columns and refinance_index in monthly_cf_df.index:
        noi_next_month = monthly_cf_df.at[refinance_index, 'Net Operating Income']
        annualized_noi = noi_next_month * 12
        annualized_debt_service = monthly_payment * 12
        refinance_information_df['Going-In DSCR'] = annualized_noi / annualized_debt_service if annualized_debt_service != 0 else 0
    else:
        print("Invalid refinance month or 'Net Operating Income' not a column")
        refinance_information_df['Going-In DSCR'] = 0
except ZeroDivisionError:
    refinance_information_df['Going-In DSCR'] = 0
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    refinance_information_df['Going-In DSCR'] = 0


# calculation of Going-In Debt Yield , refinance information:
try:
    refinance_index = refinance_month
    # Ensure refinance_month is int and exists in the columns
    if 'Net Operating Income' in monthly_cf_df.columns and refinance_index in monthly_cf_df.index:
        # Retrieving the monthly value for 'Net Operating Income' and converting it to annual
        noi_next_month = monthly_cf_df.at[refinance_index + 1, 'Net Operating Income'] * 12

        refinance_information_df['Going-In Debt Yield'] = noi_next_month / loan_amount if loan_amount != 0 else 0
    else:
        refinance_information_df['Going-In Debt Yield'] = 0
except KeyError:
    refinance_information_df['Going-In Debt Yield'] = 0
except ZeroDivisionError:
    refinance_information_df['Going-In Debt Yield'] = 0


# calculation of refinance_loan_amount
if refinance_month < hold_period:

    max_allowed_loan_amount = refinance_information_df['Max Allowed Loan Amount at 1.18x DSCR'].iloc[0]

    # Define the range of months for the following year after refinance
    start_month = refinance_month + 1
    end_month = start_month + 11

    # Filter and sum the NOI for the specified 12-month range
    if start_month in monthly_cf_df.index and end_month in monthly_cf_df.index:
        noi_12_months = monthly_cf_df.loc[start_month:end_month, 'Net Operating Income'].sum()
    else:
        noi_12_months = 0  # Default to 0 if the range is not completely within the DataFrame
else:
    noi_12_months = 0  # No refinancing if refinance_month is not before hold_period

# Step 2: Calculate the potential loan amount from NOI
potential_loan_amount = (noi_12_months / cap_rate) * LTV_refinance

# Step 3: Determine the minimum of the max allowed loan or the calculated potential loan
loan_amount_refinance = min(max_allowed_loan_amount, potential_loan_amount if noi_12_months > 0 else 0)

# Output or store the calculated loan amount for refinance
refinance_information_df['Loan Amount'] = loan_amount_refinance


# calculation of Refinance Loan Origination Fee and Initial Loan Prepayment Penalty:
refinance_information_df['Refinance Loan Origination Fee'] = -0.01 * loan_amount_refinance
refinance_information_df['Initial Loan Prepayment Penalty'] = -0.02 * loan_amount

# calculation of Total Refinance Cost
try:
    # Retrieve values from the DataFrame
    refinance_loan_origination_fee = refinance_information_df['Refinance Loan Origination Fee'].iloc[0]
    initial_loan_prepayment_penalty = refinance_information_df['Initial Loan Prepayment Penalty'].iloc[0]
    
    # Calculate the percentage
    if loan_amount_refinance != 0:
        fee_percentage = (-refinance_loan_origination_fee - initial_loan_prepayment_penalty) / loan_amount_refinance
        # Convert to string and format as percentage
        total_refinance_cost = f"{fee_percentage:.0%}"
    else:
        total_refinance_cost = 0
except Exception as e:
    total_refinance_cost = 0

# Output or store the result
refinance_information_df['Total Refinance Cost'] = total_refinance_cost


##################################!!!!!!!!!!!!!!!!!!!__________Annual CF Sheet____________!!!!!!!!!!!!!!!!###############################

# Aggregating data annually
numeric_cols = monthly_cf_df.select_dtypes(include=[np.number]).columns
annual_cf_df = monthly_cf_df[numeric_cols].groupby(monthly_cf_df['Year']).sum()

print("Index of annual_cf_df:", annual_cf_df.index)

# If specific operations need to be performed on different columns, you can use .agg() with a dictionary:
annual_cf_df = monthly_cf_df.groupby('Year').agg({
    'Gross Potential Rent': 'sum',
    'Physical Vacancy': 'sum',
    'Concessions, Bad Debt, Loss to Lease': 'sum',
    'Net Rental Income': 'sum',
    'Other Income': 'sum',
    'Total Other Income': 'sum',
    'Effective Gross Revenue': 'sum',
    'Expenses': 'sum',
    'Management Fee': 'sum',
    'Personnel': 'sum',
    'General & Administrative': 'sum',
    'Marketing': 'sum',
    'Repairs & Maintenance': 'sum',
    'Turnover': 'sum',
    'Contract Services': 'sum',
    'Utilities': 'sum',
    'Utility Reimbursements': 'sum',
    'Property Taxes': 'sum',
    'Insurance': 'sum',
    'Capital Reserves': 'sum',
    'Total Expenses': 'sum',
    'Net Operating Income': 'sum',
    'Capital & Partnership Expenses': 'sum',
    'Construction Expenses': 'sum',
    'Construction Management Fee': 'sum',
    'Asset Management Fee': 'sum',
    'Cash Flow Before Debt Service': 'sum',
    'Debt Service': 'sum',
    'Interest Payment': 'sum',
    'Principal Payment': 'sum',
    'Total Debt Service': 'sum',
    'Cash Flow After Debt Service': 'sum',
    'Acquisition & Sale Information': 'sum',
    'Purchase Price': 'sum',
    'Acquisition Fee': 'sum',
    'Closing Costs': 'sum',
    'Upfront Capex': 'sum',
    'Sale Price': 'sum',
    'Disposition Fee': 'sum',
    'Costs of Sale': 'sum',
    'Unlevered Net Cash Flow': 'sum',
    'Loan Information': 'sum',
    'Loan Funding': 'sum',
    'Loan Payoff': 'sum',
    'Loan Fees': 'sum',
    'Working Construction Capital': 'sum',
    'Working Capital Contributed': 'sum',
    'Working Capital Distributed': 'sum',
    'Levered Net Cash Flow': 'sum',
    'Ending Working Capital Balance': 'sum',
    'Net Sales/Refinance Proceeds': 'sum'
})

# from Acquisition Information in Snapsheet:
try:
    # Accessing 'Net Operating Income' for the second year
    noi_for_second_year = annual_cf_df.loc[2, 'Net Operating Income']
    print(f"Net Operating Income for the second year: {noi_for_second_year}")

    # Calculate the Going-In Cap Rate if needed
    if purchase_price != 0:
        going_in_cap_rate = noi_for_second_year / purchase_price
        acquisition_information_df['Going-In Cap Rate'] = going_in_cap_rate
        print(f"Going-In Cap Rate: {going_in_cap_rate}")
    else:
        acquisition_information_df['Going-In Cap Rate'] = 0
        print("Purchase price is zero, cannot compute Going-In Cap Rate.")

except KeyError as e:
    print(f"Data for the second year is not available: {e}")
except ZeroDivisionError:
    acquisition_information_df['Going-In Cap Rate'] = 0
    print("Purchase price is zero, adjusted to avoid division by zero.")

transposed_annual_cf_df = annual_cf_df.T  # Transpose

# Renaming the index to reflect that these are now metrics
transposed_annual_cf_df.rename_axis('Year', inplace=True)

# Resetting the index to make 'Metrics' a column, if needed
transposed_annual_cf_df.reset_index(inplace=True)

# From 'Revenue and Expenses' sheet:

#gross_potential_rent_1_year = annual_cf_df.loc[1, 'Gross Potential Rent']
vacancy_1_year = annual_cf_df.loc[1, 'Physical Vacancy']
concession_baddebt_losstolease_1_year = annual_cf_df.loc[1, 'Concessions, Bad Debt, Loss to Lease']
net_rental_income_1_year = gross_potential_rent_1_year + vacancy_1_year + concession_baddebt_losstolease_1_year 
total_other_income_1_year = annual_cf_df.loc[1, 'Total Other Income']
effective_gross_income_1_year = net_rental_income_1_year + total_other_income_1_year

year1_projected_revenue_df = pd.DataFrame({
    'Year 1 Projected Revenue': '',
    'Gross Potential Rent': [gross_potential_rent_1_year],
    'Vacancy': [vacancy_1_year],
    'Concessions, Bad Debt & Loss to Lease': [concession_baddebt_losstolease_1_year],
    'Net Rental Income': [net_rental_income_1_year],
    'Other Income': [total_other_income_1_year],
    'Effective Gross Income': [effective_gross_income_1_year]
})

projected_net_operating_income = effective_gross_income_1_year - projected_total_expenses

projected_net_operating_income_df = pd.DataFrame({
    'Projected Net Operatng Income': [projected_net_operating_income]
})


##################################!!!!!!!!!!!!!!!!!!!__________Output____________!!!!!!!!!!!!!!!!###############################

# Write the DataFrame to an Excel file with a new sheet named "Output"
output_file_path = '/Users/andriydavydyuk/Desktop/Underwriting_Project/Output.xlsx'

# Transpose the data frames to look similar as in the example
current_other_income_df = current_other_income_df.set_index('Current Other Income').T.rename_axis('Current Other Income').reset_index()
rent_increase_timing_df = rent_increase_timing_df.set_index('Rent Increase Timing').T.rename_axis('Rent Increase Timing').reset_index()
monthly_increase_data_df = monthly_increase_data_df.set_index('Month').T.rename_axis('Month').reset_index()
operating_expense_assumptions_df = operating_expense_assumptions_df.set_index('Operating Expense Assumptions').T.rename_axis('Operating Expense Assumptions').reset_index()
property_tax_growth_df = property_tax_growth_df.transpose()
property_tax_information_df = property_tax_information_df.set_index('Property Tax Information').T.rename_axis('Property Tax Information').reset_index()
hold_period_property_tax_df = hold_period_property_tax_df.set_index('Year').T.rename_axis('Year').reset_index()
t1_annulized_revenue_df = t1_annulized_revenue_df.set_index('T-1 Annulized Revenue').T.rename_axis('T-1 Annulized Revenue').reset_index()
t12_operating_expenses_df = t12_operating_expenses_df.set_index('T-12 Operating Expenses').T.rename_axis('T-12 Operating Expenses').reset_index()
projected_operating_expenses_df = projected_operating_expenses_df.set_index('Projected Operating Expenses').T.rename_axis('Projected Operating Expenses').reset_index()
utilities_df = utilities_df.set_index('Utilities').T.rename_axis('Utilities').reset_index()
year1_projected_revenue_df = year1_projected_revenue_df.set_index('Year 1 Projected Revenue').T.rename_axis('Year 1 Projected Revenue').reset_index()

refinance_information_df = refinance_information_df.set_index('Refinance Information').T.rename_axis('Refinance Information').reset_index()
property_detail_df = property_detail_df.set_index('Property Details').T.rename_axis('Property Details').reset_index()
exit_assumptions_df = exit_assumptions_df.set_index('Exit Assumptions').T.rename_axis('Exit Assumptions').reset_index()
sale_details_df = sale_details_df.set_index('Sale Details').T.rename_axis('Sale Details').reset_index()
total_uses_at_close_df = total_uses_at_close_df.set_index('Total Uses (at Close)').T.rename_axis('Total Uses (at Close)').reset_index()
sources_at_close_df = sources_at_close_df.set_index('Sources (at Close)').T.rename_axis('Sources (at Close)').reset_index()
# explicitly transposing:
acquisition_loan_information_df = acquisition_loan_information_df.T
acquisition_loan_information_df.columns = ['']
acquisition_loan_information_df.reset_index(inplace=True)
acquisition_loan_information_df.rename(columns={'index': 'Acquisition Loan Information'}, inplace=True)
# explicitly transposing:
acquisition_information_df = acquisition_information_df.T
acquisition_information_df.columns = ['']
acquisition_information_df.reset_index(inplace=True)
acquisition_information_df.rename(columns={'index': 'Acquisition Information'}, inplace=True)
# explicitly transposing:
capital_and_operating_details_df = capital_and_operating_details_df.T
capital_and_operating_details_df.columns = ['']
capital_and_operating_details_df.reset_index(inplace=True)
capital_and_operating_details_df.rename(columns={'index': 'Capital & Operating Details'}, inplace=True)

monthly_cf_df = monthly_cf_df.set_index('Month').T.rename_axis('Month').reset_index()

#annual_cf_df = annual_cf_df.set_index('Year').T.rename_axis('Year').reset_index()

# Create an ExcelWriter object
with pd.ExcelWriter(output_file_path) as writer:
    # Revenue & Expenses sheet output:
    dataframes = [
        (in_place_rents_table_df, True), # True - start from the new row, False - the same row, different column
        (stabilized_rent_df, False),
        (rent_premium_df, False),
        (comps_table_df, False),
        (current_other_income_df, True),
        (annual_growth_rates_df, True),
        (rent_increase_timing_df, True),
        (monthly_increase_data_df, True),
        (operating_expense_assumptions_df, True),
        (t1_annulized_revenue_df, True),
        (year1_projected_revenue_df, False),
        (property_tax_growth_df, False),
        (t12_operating_expenses_df, True),
        (projected_operating_expenses_df, False),
        (property_tax_information_df, False),
        (net_operating_income_df, True),
        (projected_net_operating_income_df, False),
        (utilities_df, True)

    ]

    start_row = 1
    start_col = 0
    row_spacing = 3
    col_spacing = 2

    # Initialize variables to keep track of the current row and column position
    current_row = start_row
    current_col = start_col
    row_heights = []  # List to store heights of DataFrames in the current row
    last_row_end = start_row  # This will hold the end position of special placement cases

    for idx, (df, start_new_row) in enumerate(dataframes):
        if start_new_row:
            if row_heights:  # Calculate new row position if there were DataFrames in the last row
                current_row = max(current_row, max(row_heights) + row_spacing)
                row_heights = []  # Reset for new row
            current_col = start_col  # Reset column position to start

        # Special column spacing for specific DataFrames
        if df is t1_annulized_revenue_df or df is year1_projected_revenue_df:
            col_spacing = 3  # Increase to three columns
        else:
            col_spacing = 2  # Reset back to normal after applying it once
            

        # Handle direct placement under property_tax_growth_df without added spacing
        if df is property_tax_growth_df:
            last_row_end = current_row + df.shape[0]  # Calculate end row of property_tax_growth_df

        if df is property_tax_information_df:
            current_row = last_row_end + 2 # Place directly under property_tax_growth_df

        if df is net_operating_income_df:
            current_row = max(current_row, last_row_end + row_spacing)  # Start new row properly spaced from the last special case


        df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=current_col, startrow=current_row)
        current_col += df.shape[1] + col_spacing  # Update the column position for the next DataFrame
        row_heights.append(current_row + df.shape[0])  # Store the ending row of the current DataFrame

    #start_col = in_place_rents_table_df.shape[1] + 1  

    #start_row_after_in_place_rents = in_place_rents_table_df.shape[0] + 4
    #start_row_after_current_other_income = current_other_income_df.shape[0] + 4
    #start_row_after_annual_growth_rates = start_row_after_current_other_income.shape[0] + 4

    #in_place_rents_table_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 1)
    
    '''stabilized_rent_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=start_col, startrow = 1)

    rent_premium_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=15, startrow = 1)

    comps_table_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=18, startrow = 1)

    current_other_income_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = start_row_after_in_place_rents)

    annual_growth_rates_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = start_row_after_current_other_income)

    rent_increase_timing_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = start_row_after_annual_growth_rates)

    monthly_increase_data_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 38)


    operating_expense_assumptions_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 45)
    
    property_tax_growth_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=14, startrow = 47)

    property_tax_information_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=14, startrow = 51)

    hold_period_property_tax_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=14, startrow = 60)

    t1_annulized_revenue_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 50)

    t12_operating_expenses_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 60)

    year1_projected_revenue_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=6, startrow = 50)

    projected_operating_expenses_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=6, startrow = 60)

    net_operating_income_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 75)

    utilities_df.to_excel(writer, sheet_name='Revenue & Expenses', index=False, startcol=0, startrow = 79)'''

    # Snapshot sheet output:

    property_detail_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=1, startrow = 1)
    acquisition_information_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=1, startrow = 7)
    acquisition_loan_information_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=6, startrow = 1)
    capital_and_operating_details_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=6, startrow = 11)
    sources_at_close_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=6, startrow = 17)
    total_uses_at_close_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=6, startrow = 22)
    refinance_information_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=12, startrow = 1)
    exit_assumptions_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=12, startrow = 18)
    sale_details_df.to_excel(writer, sheet_name='Snapshot', index=False, startcol=12, startrow = 26)
    

    # Monthly CF sheet output:
    
    monthly_cf_df.to_excel(writer, sheet_name='Monthly CF', index=False, startcol=0, startrow = 0)

    # Annual CF sheet output:

    transposed_annual_cf_df.to_excel(writer, sheet_name='Annual CF', index=False, startcol=0, startrow = 0)
