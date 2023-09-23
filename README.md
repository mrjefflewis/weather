# Python Project README

This repository contains a Python project that utilizes the Google Sheets API to work with a specific Google Sheets document. Follow the steps below to set up and run the project:

## Prerequisites

- **Python**: Ensure you have Python 3.x installed on your system. If not, download and install it from the [official Python website](https://www.python.org/downloads/).

## Setting Up the Project

1. **Create a Virtual Environment**:
   
   Create a new virtual environment named "venv":

   ```bash
   python -m venv venv
   ```

2. **Activate the Virtual Environment**:

   - **On Windows**:
     ```bash
     venv\Scripts\activate
     ```

   - **On Unix or MacOS**:
     ```bash
     source venv/bin/activate
     ```

3. **Install Dependencies**:
   
   Install the project dependencies using pip and the provided `requirements.txt`:

   ```bash
   pip install -r requirements.txt
   ```

4. **Google Sheets API Key**:
   
   Obtain a Google Sheets API key. Save the API key as `service_account.json` and place it in `~/.config/gspread/`.

## Configuring the Project

1. **Copy the Google Sheets Document**:
   
   Make a copy of the Google Sheets document available at the following URL:
   [https://docs.google.com/spreadsheets/d/1j_CulCc3jxLVv-pTeYnyd35H6sFlLlAJa9wIrr5XHMs](https://docs.google.com/spreadsheets/d/1j_CulCc3jxLVv-pTeYnyd35H6sFlLlAJa9wIrr5XHMs)

2. **Update the Sheets URL Parameter**:

   Open `main.py` and update the `sheets_url` parameter with the new Google Sheets document URL.

## Running the Project

1. **Run the Python Script**:
   
   Run the main Python script:

   ```bash
   python main.py
   ```

   This will execute the script and interact with the Google Sheets API.

## Additional Notes

- Ensure you have the necessary permissions and access to the Google Sheets document and the Google Sheets API for the desired operations.
- For any issues or inquiries, contact the project maintainers or contributors.

---

This should render properly on GitHub. Let me know if you need further assistance or adjustments.