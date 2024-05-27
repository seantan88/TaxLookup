import pandas as pd

# Load the sample data Excel file
sample_data = pd.read_excel('tax_sample_data.xlsx')

# Load the tax bracket and rate data Excel file
tax_data = pd.read_excel('Tax-Data.xlsx')

# Province name mapping
province_map = {
    'NF': 'Newfoundland TONI',
    'PEI': 'PEI TONI',
    'NS': 'Nova Scotia TONI',
    'NB': 'New Brunswick TONI',
    'QC': 'Quebec TONI',
    'ON': 'Ontario TONI',
    'MB': 'Manitoba TONI',
    'SK': 'Saskatchewan TONI',
    'AB': 'Alberta TONI',
    'BC': 'British Columbia TONI',
    'YT': 'Yukon TONI',
    'NT': 'Northwest territories TONI',
    'NU': 'Nunavut TONI'
}

# Function to extract brackets and rates for a given province and year with integer year columns
def extract_tax_data_for_province_year(tax_data, province, year):
    province_data = tax_data[tax_data['Province'].fillna(method='ffill') == province]
    
    brackets = []
    rates = []
    
    for i, row in province_data.iterrows():
        if 'bracket' in str(row['Variable']).lower():
            if len(brackets) > 0:
                brackets[-1]['top'] = row[year]  # Set the upper bound of the previous bracket
            brackets.append({
                'below': row[year],
                'top': 100000000  # Temporary value, will be updated by the next bracket
            })
        elif 'rate' in str(row['Variable']).lower():
            rates.append(row[year])
    
    # Remove the last bracket if it's just a placeholder with top value 0
    if brackets and brackets[-1]['top'] == 100000000:
        brackets.pop()
    
    return {
        'brackets': brackets,
        'rates': rates
    }

# Function to get tax information for a given income
def get_tax_info(income, tax_data):
    brackets = tax_data['brackets']
    rates = tax_data['rates']
    
    for i, bracket in enumerate(brackets):
        if income > bracket['below'] and (income <= bracket['top'] or bracket['top'] == 100000000):
            lower_bound = bracket['below']
            upper_bound = bracket['top']
            rate = rates[i] if i < len(rates) else None
            return rate, upper_bound, lower_bound
    
    return None, None, None

# Function to process the sample data to add the required columns
def process_sample_data(sample_data, tax_data):
    processed_data = sample_data.copy()
    processed_data['Upper bound bracket'] = None
    processed_data['Lower bound bracket'] = None
    processed_data['Nearest bracket to income'] = None
    processed_data['Distance to nearest bracket'] = None
    processed_data['Current marginal tax rate'] = None
    
    for index, row in processed_data.iterrows():
        income = row['Income']
        
        rate, upper_bound, lower_bound = get_tax_info(income, tax_data)
        
        if rate is not None:
            processed_data.at[index, 'Upper bound bracket'] = upper_bound
            processed_data.at[index, 'Lower bound bracket'] = lower_bound
            processed_data.at[index, 'Current marginal tax rate'] = rate
            
            # Calculate the nearest bracket and distance
            distance_to_upper = abs(upper_bound - income)
            distance_to_lower = abs(income - lower_bound)
            
            if distance_to_upper < distance_to_lower:
                processed_data.at[index, 'Nearest bracket to income'] = upper_bound
                processed_data.at[index, 'Distance to nearest bracket'] = distance_to_upper
            else:
                processed_data.at[index, 'Nearest bracket to income'] = lower_bound
                processed_data.at[index, 'Distance to nearest bracket'] = distance_to_lower
    
    return processed_data

# Prompt for province and year
province_input = input("Enter the province (e.g., AB for Alberta): ")
year = int(input("Enter the year: "))

# Get the full province name
province_full = province_map.get(province_input.upper())

if province_full:
    # Extract tax data for the given province and year
    tax_data_for_province_year = extract_tax_data_for_province_year(tax_data, province_full, year)
    
    # Process the sample data with the extracted tax data
    processed_data = process_sample_data(sample_data, tax_data_for_province_year)
    
    # Save the processed data to a new Excel file
    processed_data.to_excel(f"processed_tax_data_{province_input.lower()}_{year}.xlsx", index=False)
    
    print(f"Processed data saved to processed_tax_data_{province_input.lower()}_{year}.xlsx")
else:
    print(f"Invalid province code: {province_input}")
