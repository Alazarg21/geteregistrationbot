# GETE Registration Bot 🎓

A Python-powered Telegram bot for student registration and record management. It collects student information, stores records in an Excel workbook, and allows authorized administrators to search, edit, delete, export, and manage registrations.

## Features

### Student Registration
- Welcome message and step-by-step registration.
- Collects full name, phone number, and current address.
- Validates submitted information.
- Requests confirmation before saving.
- Automatically generates a Student ID.
- Records the registration date and time.
- Prevents duplicate registrations from the same Telegram account.

### Administrator Management
- View all registered students.
- Search records by name, phone number, address, or Student ID.
- Edit a student's information using their full name or Student ID.
- Delete student records with confirmation.
- Export the Excel workbook through Telegram.
- Broadcast announcements to registered students.
- Restrict administrator commands to authorized Telegram user IDs.

### Excel Storage
Student data is stored in `students.xlsx`.

| Column | Description |
|---|---|
| Student ID | Unique student record identifier |
| Full Name | Student's full name |
| Phone Number | Contact number |
| Current Address | Student's address |
| Registration Date | Date and time of registration |
| Telegram User ID | Telegram account ID |
| Telegram Chat ID | Chat ID used for announcements |

The bot creates the workbook automatically if it does not exist. It also adds missing Telegram columns to a compatible existing workbook without intentionally deleting existing student records.

## Project Structure

```text
geteregistration_bot/
├── bot.py
├── README.md
├── requirements.txt
├── .env
├── .gitignore
└── students.xlsx
```

`students.xlsx` is generated automatically when the bot starts if it does not already exist.

## Requirements

- Python 3.10 or newer recommended.
- A Telegram account.
- A Telegram bot token from BotFather.
- An internet connection.

## Installation on Windows

### 1. Install Python

Download Python from:

https://www.python.org/downloads/

During installation, select **Add Python to PATH**.

Verify the installation in PowerShell:

```powershell
python --version
```

If necessary, try:

```powershell
py --version
```

### 2. Open the Project Folder

```powershell
cd D:\geteregistration_bot
```

### 3. Create a Virtual Environment

```powershell
py -m venv venv
```

Activate it:

```powershell
.\venv\Scripts\Activate.ps1
```

If PowerShell cannot find `py`, use `python` after ensuring Python is installed correctly.

### 4. Create `requirements.txt`

Add the following dependencies:

```text
python-telegram-bot>=20.0,<23.0
openpyxl>=3.1.0
python-dotenv>=1.0.0
```

Install them:

```powershell
python -m pip install -r requirements.txt
```

## Configure the Bot

### 1. Create a Telegram Bot

1. Open [@BotFather](https://t.me/BotFather).
2. Send `/newbot`.
3. Enter your bot's display name.
4. Choose a unique username ending in `bot`.
5. Copy the token provided by BotFather.

### 2. Configure Administrator Access

Find your numeric Telegram user ID using a trusted method.

Create a `.env` file in the same directory as `bot.py`:

```env
BOT_TOKEN=YOUR_TELEGRAM_BOT_TOKEN
ADMIN_ID=YOUR_TELEGRAM_USER_ID
```

Replace both placeholders with your actual values.

To configure multiple administrators instead, use:

```env
BOT_TOKEN=YOUR_TELEGRAM_BOT_TOKEN
ADMIN_IDS=123456789,987654321
```

`ADMIN_IDS` takes precedence over `ADMIN_ID` when both are configured.

**Security:** Never share your bot token or publish your `.env` file in a public repository.

## Run the Bot

Activate the virtual environment and run:

```powershell
python bot.py
```

The terminal should display the bot name, Excel file location, configured administrator count, and running status.

Open your bot in Telegram and send:

```text
/start
```

The bot will welcome the student and begin registration.

## Student Commands

| Command | Purpose |
|---|---|
| `/start` | Welcome message and registration |
| `/register` | Begin registration |
| `/help` | Display available commands |
| `/cancel` | Cancel the current operation |

### Registration Steps

1. Enter your full name.
2. Enter your phone number.
3. Enter your current address.
4. Review your information.
5. Reply `YES` to save or `NO` to cancel.

After successful registration, the bot displays the Student ID and saved details.

## Administrator Commands

Administrator commands must be used by an authorized administrator in a private chat with the bot.

| Command | Purpose |
|---|---|
| `/admin` | Display administrator controls |
| `/students` | List all registered students |
| `/search WORD` | Search student records |
| `/edit FULL_NAME` | Edit a student's details |
| `/delete FULL_NAME` | Find and delete a student with confirmation |
| `/export` | Download an Excel export |
| `/broadcast` | Send an announcement |

### Search Students

Search by full name:

```text
/search Abebe Kebede
```

Search by part of a name:

```text
/search Abebe
```

You can also search by Student ID, phone number, or address.

### Edit Student Information

Find a student by their full name:

```text
/edit Abebe Kebede
```

Or use the Student ID:

```text
/edit 1
```

Choose the field when prompted:

- `name`
- `phone`
- `address`

Then enter the new value.

If multiple students have the same full name, use the Student ID to identify the correct record.

### Delete a Student

Find a student by full name:

```text
/delete Abebe Kebede
```

Or by Student ID:

```text
/delete 1
```

The bot displays the selected student's information and requests confirmation. Reply `YES` to delete or `NO` to cancel.

Deletion permanently removes the selected row from the Excel workbook. Back up your workbook regularly.

### Export Student Records

Send:

```text
/export
```

The bot creates a temporary copy of `students.xlsx` and sends it to the administrator as an Excel document.

### Broadcast Announcements

1. Send `/broadcast`.
2. Enter the announcement text.
3. Review the preview and recipient count.
4. Reply `YES` to send or `NO` to cancel.
5. Review the delivery report.

Recipients must have registered with the bot and have saved chat IDs. Delivery may fail if a student has blocked the bot or cannot receive messages.

## Excel File Management

The bot stores records in `students.xlsx` beside `bot.py`.

Recommended practices:

- Back up the workbook regularly.
- Keep the workbook in the project directory.
- Close the workbook in Excel before the bot writes to it.
- Do not rename or remove required column headers.
- Restrict access because the workbook contains personal information.
- Keep backups in a secure location.

The workbook is a local file, not a cloud database. The bot must be running to process Telegram updates.

## Troubleshooting

### Python Was Not Found

Install Python from https://www.python.org/downloads/, enable PATH support, reopen PowerShell, and check:

```powershell
py --version
```

### Missing BOT_TOKEN or ADMIN_ID

Check that `.env` is in the same directory as `bot.py`.

Ensure the filename is `.env`, not `.env.txt`, and that the values are correct.

### Module Not Found

Activate your virtual environment and reinstall the dependencies:

```powershell
.\venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

### Excel Permission Error

Close `students.xlsx` in Excel and retry the operation. Check that the project folder is writable.

### Bot Does Not Respond

- Confirm that the bot token is correct.
- Check the internet connection.
- Make sure `python bot.py` is still running.
- Review the terminal for error messages.
- Verify that you are messaging the correct Telegram bot.

### Broadcast Delivery Failures

Check the broadcast report. Confirm that students have registered with the bot and that their chat IDs are saved.

## Limitations

- Students cannot edit their own records in this version; administrators manage edits.
- Search by full name is case-insensitive. If names are duplicated, use Student ID for editing or deleting.
- Excel is suitable for a small registration system but does not provide robust concurrent database access.
- The bot does not include automatic cloud backups.
- Telegram usernames are not stored by this version.

## Future Improvements

Possible enhancements include:

- Student self-service record updates.
- Education, department, and skills fields.
- CV uploads.
- Administrator approval workflows.
- Automated backups.
- Migration to SQLite or PostgreSQL.
- Deployment to a server for continuous availability.

## License

Specify a license before distributing or publishing this project.

---

**GETE Registration Bot** — Student registration and administration through Telegram.
