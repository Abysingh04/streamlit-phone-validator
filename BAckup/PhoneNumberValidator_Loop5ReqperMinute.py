import requests
import pandas as pd
import time
import os

# File paths
input_file = "D:/Python/Streamlit/Framework/InputPhoneNumber.xlsx"
output_file = "output.xlsx"

# Check if file exists
if not os.path.exists(input_file):
    print(f"❌ Error: Input file '{input_file}' not found!")
    exit()

# Read Excel file as text (to prevent formatting issues)
try:
    df = pd.read_excel(input_file, dtype=str)  # Load sheet automatically
    df.columns = df.columns.str.strip()  # Remove extra spaces
except Exception as e:
    print(f"❌ Error loading Excel file: {e}")
    exit()

# Ensure "Phone_Number" column exists
if "Phone_Number" not in df.columns:
    print(f"❌ Error: 'Phone_Number' column is missing! Available columns: {df.columns}")
    exit()

print("🚀 Starting phone number validation process...")

# API details
api_base_url = "http://phone-number-api.com/csv/"
fields = "status,message,numberType,numberValid,isDisposible,numberCountryCode,numberAreaCode,formatE164,formatInternational,carrier,continent,countryName,country,region,regionName,city,zip,query"

# Expected Response Headers
expected_headers = [
    "Status", "Message", "Number Type", "Number Valid", "Is Disposable",
    "Country Code", "Area Code", "E164 Format", "International Format",
    "Carrier", "Continent", "Country Name", "Country", "Region",
    "Region Name", "City", "ZIP", "Query"
]

# Results List
results = []

# Loop through each phone number
for index, phone_number in enumerate(df["Phone_Number"]):
    if pd.isna(phone_number) or phone_number.strip() == "":
        print("⚠️ Skipping empty phone number.")
        continue

    phone_number = phone_number.strip()  # Remove spaces

    # Ensure the phone number starts with "+"
    if not phone_number.startswith("+"):
        print(f"⚠️ Warning: '{phone_number}' is missing '+'. Skipping...")
        continue

    # Construct API Request URL
    url = f"{api_base_url}?number={phone_number}&fields={fields}"
    print(f"📡 Sending request for: {phone_number} → {url}")

    try:
        # Send GET Request
        response = requests.get(url, timeout=10)

        # Check if successful
        if response.status_code == 200:
            values = response.text.strip().split(",")

            # Handle Column Mismatches
            if len(values) < len(expected_headers):
                values += ["N/A"] * (len(expected_headers) - len(values))  # Fill missing fields
            elif len(values) > len(expected_headers):
                print(f"⚠️ Extra fields in API response for {phone_number}. Adjusting...")
                expected_headers = expected_headers[:len(values)]  # Trim headers

            # Store Result
            results.append(values)
        else:
            print(f"❌ API Error {response.status_code} for {phone_number}: {response.text}")

        # ⏳ Add Delay (12-15 sec) to Maintain 5 Requests per Minute
        time.sleep(12)

    except requests.exceptions.RequestException as e:
        print(f"⚠️ Request failed for {phone_number}: {e}")

# Save Results to Excel
if results:
    output_df = pd.DataFrame(results, columns=expected_headers)
    output_df.to_excel(output_file, index=False)
    print(f"\n✅ Results saved to {output_file}")
else:
    print("❌ No valid results to save.")
