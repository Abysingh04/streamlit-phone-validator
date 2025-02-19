import requests

# API endpoint for temporary phone number validation
url = "https://phone-number-api.com/api/v1/temporary-validation"

# Sample phone numbers (you can modify this list)
params = {
    "numbers": "1234567890,9876543210"  # Replace with actual numbers
}

try:
    response = requests.get(url, params=params)

    if response.status_code == 200:
        print("Response JSON:", response.json())  # Print response data
    else:
        print("Error:", response.status_code, response.text)

except requests.ex
