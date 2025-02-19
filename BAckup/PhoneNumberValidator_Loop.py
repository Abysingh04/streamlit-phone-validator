import requests
import pandas as pd
import time
import random

# File paths
input_file = "D:/Python/Streamlit/Framework/InputPhoneNumber.xlsx"  # Input Excel file
output_file = "output.xlsx"  # Output file for results

# Read Excel file as text (to prevent formatting issues)
try:
    df = pd.read_excel(input_file, dtype=str, sheet_name="Sheet2")  # Treat all data as text
except FileNotFoundError:
    print(f"❌ Error: File '{input_file}' not found!")
    exit()

# Ensure "Phone_Number" column exists
if "Phone_Number" not in df.columns:
    print("❌ Error: 'Phone_Number' column is missing in the input file.")
    exit()

# API details
api_base_url = "http://phone-number-api.com/csv/"
fields = "status,message,numberType,numberValid,isDisposible,numberCountryCode,numberAreaCode,formatE164,formatInternational,carrier,continent,countryName,country,region,regionName,city,zip,query"

# **Expected Response Headers**
expected_headers = [
    "Status", "Message", "Number Type", "Number Valid", "Is Disposable",
    "Country Code", "Area Code", "E164 Format", "International Format",
    "Carrier", "Continent", "Country Name", "Country", "Region",
    "Region Name", "City", "ZIP", "Query"
]

# **Results List**
results = []

# **Loop through each phone number**
for index, phone_number in enumerate(df["Phone_Number"]):
    if pd.isna(phone_number) or phone_number.strip() == "":
        print("⚠️ Skipping empty phone number.")
        continue

    phone_number = phone_number.strip()  # Remove spaces

    # Ensure the phone number starts with "+"
    if not phone_number.startswith("+"):
        print(f"⚠️ Warning: '{phone_number}' is missing '+'. Skipping...")
        continue

    # **Construct API Request URL**
    url = f"{api_base_url}?number={phone_number}&fields={fields}"
    
    try:
        # **Send GET Request**
        response = requests.get(url, timeout=10)

        # **Check if successful**
        if response.status_code == 200:
            values = response.text.strip().split(",")

            # **Handle Column Mismatches**
            if len(values) < len(expected_headers):
                values += ["N/A"] * (len(expected_headers) - len(values))  # Fill missing fields
            elif len(values) > len(expected_headers):
                print(f"⚠️ Extra fields in API response for {phone_number}. Adjusting...")
                expected_headers = expected_headers[:len(values)]  # Trim headers

            # **Store Result**
            results.append(values)
        else:
            print(f"❌ API Error {response.status_code} for {phone_number}: {response.text}")
            
        # **⏳ Add Delay (12 sec) to Maintain 5 Requests per Minute**
        time.sleep(12)

        # # **⏳ Add Delay (Random 2-5 sec) to Prevent Ban**
        # time.sleep(random.uniform(1, 2))

    except requests.exceptions.RequestException as e:
        print(f"⚠️ Request failed for {phone_number}: {e}")

# **Save Results to Excel**
if results:
    output_df = pd.DataFrame(results, columns=expected_headers)
    output_df.to_excel(output_file, index=False)
    print(f"\n✅ Results saved to {output_file}")
else:
    print("❌ No valid results to save.")
