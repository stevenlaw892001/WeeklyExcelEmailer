# Weekly Report Emailer

## Overview

This Python script automates sending weekly CPU and disk utilization reports via email. It attaches `.xlsx` files from a specified directory, covering the previous week's date range (Monday to Friday). The script dynamically generates date ranges, filters relevant files, and sends emails using a configured SMTP server.

**Key Features**:
- Dynamically calculates the previous week's date range.
- Attaches `.xlsx` reports matching the date range (e.g., `24Mar2025-28Mar2025`).
- Sends emails via an SMTP server with SSL encryption.
- Includes error handling for missing files and SMTP failures.

## Prerequisites

- Python 3.6 or higher.
- Required Python package:
  ```
  pip install python-dotenv
  ```
- Access to an SMTP server (configured via environment variables).
- `.xlsx` reports stored in a specified directory.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/weekly-report-emailer.git
   ```
2. Navigate to the project directory:
   ```bash
   cd weekly-report-emailer
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   (Or install `python-dotenv` directly: `pip install python-dotenv`)

## Configuration

1. Create a `.env` file in the project root with the following variables:
   ```
   SENDER_EMAIL=sender@example.com
   RECEIVER_EMAIL=receiver@example.com
   SMTP_SERVER=smtp.example.com
   SMTP_PORT=465
   REPORTS_DIR=path/to/reports
   ```
   See `.env.example` for a template.
2. Ensure `.xlsx` reports are in the `REPORTS_DIR` directory, named with the date range (e.g., `Report_24Mar2025-28Mar2025.xlsx`).

**Note**: The `.env` file is included in `.gitignore` to prevent sensitive data from being uploaded.

## Usage

1. Place `.xlsx` reports in the configured directory.
2. Run the script:
   ```bash
   python send_weekly_report.py
   ```
3. The script will:
   - Identify the previous week's date range.
   - Attach matching `.xlsx` files.
   - Send the email via the configured SMTP server.
   - Display "Email sent successfully!" or an error message.

## Project Structure

```
├── send_weekly_report.py  # Main script
├── .gitignore             # Ignores sensitive files
├── README.md              # Project documentation
├── requirements.txt       # Dependency list
├── LICENSE                # MIT License
├── CODE_OF_CONDUCT.md     # Community guidelines
├── CONTRIBUTING.md        # Contribution guidelines
├── .env.example           # Environment variable template
├── .github/
│   ├── ISSUE_TEMPLATE/    # Issue templates
│   └── PULL_REQUEST_TEMPLATE.md  # PR template
```

## Notes

- Ensure `.xlsx` files exist in the configured directory, or the script will exit with a "No files found" error.
- The script uses SSL with port 465 by default. Modify `send_weekly_report.py` if your SMTP server uses a different configuration.
- For SMTP servers requiring authentication, add `server.login()` logic as needed.

## Contributing

Want to contribute? Please read our [Contributing Guidelines](CONTRIBUTING.md) for details on submitting issues and pull requests. Kindly review our [Code of Conduct](CODE_OF_CONDUCT.md) to ensure respectful and inclusive interactions.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For questions or feedback, please open a GitHub Issue or contact me via GitHub (@your-username).
