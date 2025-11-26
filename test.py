import requests
import pandas as pd
import openpyxl
from datetime import datetime

# API endpoint
url = "https://gapi.cenhomes.vn/g-analyst-graphql/v1"

# Note: This API likely requires parameters such as property ID or other filters.
# You may need to inspect the website's network requests to determine required parameters.
# For example, it might need something like: params = {'propertyId': 'some_id'}
# Uncomment and modify the params dict below if needed.

# params = {
#     'propertyId': 'example_id',  # Replace with actual property ID
#     # Add other required parameters here
# }

try:
    # Make GET request to the API
    response = requests.get(url)  # Add params=params if needed
    response.raise_for_status()  # Raise an error for bad status codes

    # Parse JSON response
    data = response.json()

    # Assuming the response is a list of dictionaries or can be converted to DataFrame
    # If the structure is different, adjust accordingly
    df = pd.DataFrame(data)

    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"pricing_history_{timestamp}.xlsx"

    # Export to Excel
    df.to_excel(filename, index=False)

    print(f"Data exported successfully to {filename}")

except requests.exceptions.RequestException as e:
    print(f"Error fetching data: {e}")
except ValueError as e:
    print(f"Error parsing JSON: {e}")
except Exception as e:
    print(f"An error occurred: {e}")