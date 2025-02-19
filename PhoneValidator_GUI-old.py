import streamlit as st
import pandas as pd
import time
import requests
import io

# 🌐 API Details
API_BASE_URL = "http://phone-number-api.com/csv/"
FIELDS = "status,numberType,numberValid,numberValidForRegion,isDisposible,numberCountryCode,numberAreaCode,formatE164,formatNational,formatInternational,carrier,continent,continentCode,countryName,country,region,regionName,city,zip,offset,currency,query"

# Expected API Response Headers
EXPECTED_HEADERS = [
    "Status", "Number Type", "Number Valid", "numberValidForRegion", "Is Disposable",
    "Country Code", "Area Code", "E164 Format", "National Format", "International Format",
    "Carrier", "Continent", "Continent Code", "Country Name", "Country", "Region",
    "Region Name", "City", "ZIP", "Offset", "Currency", "Query"
]

# Streamlit UI
st.title("📞 Phone Number Validator")

uploaded_file = st.file_uploader("Upload an Excel file with phone numbers", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = df.columns.str.strip()
    
    if "Phone_Number" not in df.columns:
        st.error("❌ 'Phone_Number' column is missing! Please upload a valid file.")
    else:
        st.success("✅ File uploaded successfully!")
        
        results = []
        progress_bar = st.progress(0)
        total_numbers = len(df)
        
        for index, phone_number in enumerate(df["Phone_Number"].dropna().astype(str), start=1):
            phone_number = phone_number.strip().replace(" ", "").replace("-", "")
            
            if not phone_number.startswith("+"):
                results.append(["INVALID_FORMAT"] + ["N/A"] * (len(EXPECTED_HEADERS) - 2) + [phone_number])
                continue
            
            url = f"{API_BASE_URL}?number={phone_number}&fields={FIELDS}"
            
            try:
                response = requests.get(url, timeout=10)
                if response.status_code == 200:
                    values = response.text.strip().split(",")
                    values += ["N/A"] * (len(EXPECTED_HEADERS) - len(values))
                else:
                    values = ["API_ERROR"] + ["N/A"] * (len(EXPECTED_HEADERS) - 2) + [phone_number]
            except:
                values = ["REQUEST_FAILED"] + ["N/A"] * (len(EXPECTED_HEADERS) - 2) + [phone_number]
                
            results.append(values)
            progress_bar.progress(index / total_numbers)
            time.sleep(1)  # Simulate API rate limit
        
        st.success("✅ Processing complete!")
        result_df = pd.DataFrame(results, columns=EXPECTED_HEADERS)
        
        # Convert DataFrame to Excel for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            result_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Download Processed File",
            data=output.getvalue(),
            file_name="validated_numbers.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def main():
    
  
    st.write("**Important Instruction**: Your File must have Column called **< Phone_Number >** to work this code")

if __name__ == "__main__":
    main()
