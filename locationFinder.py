import pandas as pd
import googlemaps

def get_lat_long(api_key, address):
    gmaps = googlemaps.Client(key=api_key)
    geocode_result = gmaps.geocode(address)
    if geocode_result:
        lat = geocode_result[0]['geometry']['location']['lat']
        lng = geocode_result[0]['geometry']['location']['lng']
        return lat, lng
    else:
        return None, None

def process_excel(input_file, output_file, api_key):
    df = pd.read_excel(input_file)
    if 'Location' not in df.columns:
        raise ValueError("Excel file must contain a column named 'Location'")
    
    latitudes = []
    longitudes = []
    
    for location in df['Location']:
        lat, lng = get_lat_long(api_key, location)
        latitudes.append(lat)
        longitudes.append(lng)
    
    df['Latitude'] = latitudes
    df['Longitude'] = longitudes
    
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# Usage
API_KEY = "YOUR_GOOGLE_MAPS_API_KEY"  # Replace with your API key
INPUT_FILE = "input.xlsx"  # Replace with your input Excel file
OUTPUT_FILE = "output.xlsx"  # this will be the output file

process_excel(INPUT_FILE, OUTPUT_FILE, API_KEY)

