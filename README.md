# Mass Email Sender

This project provides two versions of a mass email sender:

1. **Version 1.0**: Uses Python's `smtplib` to send emails directly through Gmail.
2. **Version 2.0**: Uses the Gmail API to send emails, allowing for more secure and scalable email sending.

## Features

- Send personalized emails to a list of recipients.
- Read recipient details from an Excel file.
- Customize email content with recipient-specific details.

## Setup Instructions

### Prerequisites

- Python 3.7 or higher
- A Google account with Gmail enabled
- The following Python packages:

  - `pandas`
  - `openpyxl`
  - `google-auth`
  - `google-auth-oauthlib`
  - `google-auth-httplib2`
  - `google-api-python-client`

### Installation

1. Clone the repository to your local machine:

    ```bash
    git clone https://github.com/your-repo/mass-email-sender.git
    cd mass-email-sender
    ```

2. Install the required packages:

    ```bash
    pip install -r requirements.txt
    ```

### Setting Up Version 1.0

1. Open `mass_email_sender_v1.0.py`.
2. Update the following lines with your Gmail credentials:

    ```python
    sender_email = 'your-email@gmail.com'
    app_password = 'your-app-password'
    ```

3. Update the path to your Excel file:

    ```python
    excel_file_path = 'path/to/your/excel_file.xlsx'
    ```

4. Run the script:

    ```bash
    python mass_email_sender_v1.0.py
    ```

### Setting Up Version 2.0

1. Obtain OAuth 2.0 credentials:

    - Go to the [Google Cloud Console](https://console.cloud.google.com/).
    - Create a new project and enable the Gmail API.
    - Create OAuth 2.0 credentials and download the `credentials.json` file.
    - Place `credentials.json` in the project directory.

2. Run the script:

    ```bash
    python mass_email_sender_v2.0.py
    ```

3. The first time you run the script, you will be prompted to authorize the application. Follow the instructions to complete the authorization.

4. Update the path to your Excel file:

    ```python
    excel_file_path = 'path/to/your/excel_file.xlsx'
    ```

5. Run the script again:

    ```bash
    python mass_email_sender_v2.0.py
    ```

### Excel File Structure

The Excel file should have the following structure:

| Name       | Email            |
|------------|------------------|
| Recipient1 | email1@example.com |
| Recipient2 | email2@example.com |

### Notes

- For Version 1.0, you need to enable "Less secure apps" in your Google account settings or generate an App Password.
- For Version 2.0, you don't need to enable "Less secure apps" since it uses OAuth 2.0 for authentication.

## Contributing

If you want to contribute to this project, feel free to submit a pull request.
