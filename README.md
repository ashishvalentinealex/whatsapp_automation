# WhatsApp Message Automation

## Overview
This project automates sending personalized WhatsApp messages to contacts listed in an Excel sheet using Python and Selenium. It searches for each contact by name, personalizes the message, and handles exceptions for missing or inaccessible contacts.

## Features
- Automated message sending on WhatsApp Web.
- Reads contact names from an Excel sheet.
- Skips missing or inaccessible contacts gracefully.
- Sends personalized messages by dynamically inserting contact names.

## Requirements
### Software Dependencies
- Python 3.x
- Selenium WebDriver
- Chrome WebDriver
- openpyxl

### Setup
1. Install Python 3.x from [python.org](https://www.python.org/).
2. Install the required Python libraries:
   ```bash
   pip install selenium openpyxl
   ```
3. Download Chrome WebDriver compatible with your version of Chrome from [chromedriver.chromium.org](https://chromedriver.chromium.org/).
4. Ensure the Excel file (`contacts.xlsx`) contains the following:
   - The first row as a header.
   - Names of the contacts in the first column.

### Directory Structure
```
project/
|-- contacts.xlsx  # Excel sheet with contact names
|-- main.py        # Main script
|-- README.md      # This file
```

## How to Run
1. Clone the repository:
   ```bash
   git clone <repository_url>
   cd <repository_name>
   ```
2. Place your `contacts.xlsx` file in the project directory.
3. Run the script:
   ```bash
   python main.py
   ```
4. When prompted, scan the QR code to log into WhatsApp Web.
5. Watch the automation send messages!

## Customization
### Message Template
The message template can be updated by modifying this line in the script:
```python
common_message = "Hello {name}, this is an automated message from our system!"
```
Replace the content within the quotes as per your requirements.

### Excel Sheet
Ensure the first column of the Excel sheet contains the contact names. Modify the column index if your structure is different:
```python
name = row[0]  # Adjust index if needed
```

## Error Handling
- **Missing Contacts**:
  If a contact is not found, the script logs the event and skips to the next contact.
- **General Errors**:
  Any unexpected errors during execution are logged with details.

## Important Notes
- **Privacy and Permissions**:
  Ensure you have permission to send messages to the listed contacts.
- **WebDriver Compatibility**:
  Use a ChromeDriver version that matches your installed Chrome browser version.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Contributions
Contributions are welcome! Feel free to fork the repository and submit pull requests.

## Contact
For issues or feature requests, please raise an issue in the repository or contact ashishvalentinealex@gmail.com

