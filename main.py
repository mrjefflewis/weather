import gspread
import requests
from datetime import datetime

sheets_url = 'https://docs.google.com/spreadsheets/d/1E1yfsNt8MQkf5U-otmnFLwD_LXU-rMNfqiiTUvdnsDY'
read_sheets_range = "A1:B100"

weather_url = 'https://api.open-meteo.com/v1/forecast'

def lookup_weather(lat, long):
    """
    Returns the weather at the given latitude and longitude.

    Args:
        lat (float): The latitude of the location.
        long (float): The longitude of the location.

    Returns:
        dict: A dictionary containing the weather information.
    """
    params = {'latitude': lat, 'longitude': long, 'current_weather': 'true'}
    response = requests.get(weather_url, params=params)
    return response.json()

def get_sheets_data(sheets_url, sheets_range):
    """
    Returns the data from the specified Google Sheets range.

    Args:
        sheets_url (str): The URL of the Google Sheets document.
        sheets_range (str): The range of cells to retrieve.

    Returns:
        list: A list of lists containing the data from the specified range.
    """
    # Create a Google Sheets API client.
    gc = gspread.service_account()

    # Open the spreadsheet.
    sh = gc.open_by_url(sheets_url)

    # Get the worksheet.
    worksheet = sh.sheet1

    values_list = worksheet.get(sheets_range)

    return values_list

def lookup_values(values_list):
    """
    Returns a list of dictionaries containing the weather information for each location in the specified list.

    Args:
        values_list (list): A list of lists containing the latitude and longitude of each location.

    Returns:
        list: A list of dictionaries containing the weather information for each location.
    """
    latlong_values = []
    latcol, longcol = 0, 1
    row = 2

    while True:
        try:
            lat_value = values_list[row][latcol]
            long_value =values_list[row][longcol]
            weather_data = lookup_weather(lat_value, long_value)
            temperature = round(weather_data['current_weather']['temperature'])
            windspeed = weather_data['current_weather']['windspeed']
            time_str = weather_data['current_weather']['time']
            time = datetime.strptime(time_str, '%Y-%m-%dT%H:%M').strftime('%Y-%m-%d %I:%M %p')
        except IndexError:
            break

        latlong_values.append({'lat': lat_value, 'long': long_value, 'temperature': f'{temperature} F', 'windspeed': f'{windspeed} mph', 'time': time})
        row += 1
        print(row)
    
    return latlong_values

def write_values(values_list):
    """
    Writes the weather information for each location to the specified Google Sheets document.

    Args:
        values_list (list): A list of dictionaries containing the weather information for each location.
    """
    # Create a Google Sheets API client.
    gc = gspread.service_account()

    # Open the spreadsheet.
    sh = gc.open_by_url(sheets_url)

    # Get the worksheet.
    worksheet = sh.sheet1

    update_values = []

    max_row = len(values_list) + 2
    update_sheets_range = f"C3:E{max_row}"

    for row in values_list:
        update_values.append([row['temperature'], row['windspeed'], row['time']])

    worksheet.format(update_sheets_range, {"horizontalAlignment": "LEFT"})
    worksheet.format(f"E3:E{max_row}", {"horizontalAlignment": "RIGHT"})
    worksheet.update(update_sheets_range, update_values)

def main():
    """
    Retrieves the weather information for each location in the specified Google Sheets document and writes it to the document.
    """
    values_list = lookup_values(get_sheets_data(sheets_url, read_sheets_range))
    print(values_list)

    write_values(values_list)


#initialize the main function
if __name__ == '__main__':
    main()