import requests
import pandas as pd

# Example phone number
phone_number = "+17202764654"

# API URL
url = f"http://phone-number-api.com/csv/?number={phone_number}&fields=status,message,numberType,numberValid,isDisposible,numberCountryCode,numberAreaCode,formatE164,formatInternational,carrier,continent,countryName,country,region,regionName,city,zip,query"

try:
    # Send GET request
    response = requests.get(url)

    # Print raw response (for debugging)
    print("Raw Response:", response.text)

    # Check if request is successful
    if response.status_code == 200:
        # Convert CSV response to a list (only values)
        values = response.text.strip().split(",")

        # Manually define the correct column headers (must match API fields)
        headers = [
            "Status", "Number Type", "Number Valid", "Is Disposable",
            "Country Code", "Area Code", "E164 Format", "International Format",
            "Carrier", "Continent", "Country Name", "Country", "Region",
            "Region Name", "City", "ZIP", "Query"
        ]

        # Ensure headers and values have the same length
        if len(headers) != len(values):
            print(f"⚠️ Column mismatch! Headers: {len(headers)}, Values: {len(values)}")
            exit()

        # Convert to DataFrame
        df = pd.DataFrame([values], columns=headers)

        # Save to Excel
        output_file = "output.xlsx"
        df.to_excel(output_file, index=False)

        print(f"\n✅ Results saved to {output_file}")

    else:
        print(f"❌ Error {response.status_code}: {response.text}")

except requests.exceptions.RequestException as e:
    print("⚠️ Request failed:", e)
