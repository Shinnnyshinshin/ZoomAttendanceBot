# Zoom Attendance Report Generator

Automated Python scripts to generate attendance reports from Zoom meetings with PST timezone support. Available in both interactive and automated (CRON) versions.

## 📋 Features

- **Clean Reports**: Generates Excel files with only essential data (Meeting Date, Meeting Time, Participant Name)
- **Email Integration**: Automatically emails reports with participant names listed in the email body
- **Timezone Support**: Converts UTC times to PST/PDT automatically
- **Deduplication**: Combines multiple sessions for the same participant
- **Flexible Time Ranges**: Support for hours (`2h`), minutes (`30m`), or days (`1d`)
- **Two Versions**: Interactive script for manual use, CRON script for automation

## 🔧 Requirements

### Python Dependencies
```bash
pip install pandas openpyxl requests pytz python-dotenv
```

### Zoom API Credentials
1. Go to [Zoom Marketplace](https://marketplace.zoom.us/)
2. Create a Server-to-Server OAuth app
3. Get your Account ID, Client ID, and Client Secret

### Email Configuration
- SMTP server access (Gmail, Outlook, etc.)
- For Gmail: Use App Passwords, not your regular password

## ⚙️ Setup

### 1. Environment Configuration

Create a `.env` file in the same directory as your scripts:

```bash
# Zoom API Configuration
ZOOM_ACCOUNT_ID=your_account_id_here
ZOOM_CLIENT_ID=your_client_id_here
ZOOM_CLIENT_SECRET=your_client_secret_here

# Email Configuration
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_app_password_here
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Recipients (comma-separated)
EMAIL_RECIPIENTS=manager@company.com,hr@company.com
```

### 2. Script Configuration (CRON version only)

Edit the configuration variables at the top of `zoom_attendance_cron.py`:

```python
# CONFIGURATION - Update these values
MEETING_ID = "123456789"  # Your Zoom meeting ID (as string)
TIME_RANGE = "24h"        # How far back to look (2h, 30m, 1d, etc.)
SEND_EMAIL = True         # Set to False to only generate files
```

## 🚀 Usage

### Interactive Version (`zoom.py`)

Run the script and follow the prompts:

```bash
python3 zoom.py
```

**Menu Options:**
1. **Generate report** - Create attendance report with prompts for meeting ID and time range
2. **Test email** - Verify your email configuration
3. **Create .env template** - Generate a template .env file
4. **Show configuration** - Display current configuration status

### CRON Version (`zoom_attendance_cron.py`)

**Manual test run:**
```bash
python3 zoom_attendance_cron.py test
```

**Automated run (for CRON):**
```bash
python3 zoom_attendance_cron.py
```

## 📅 Setting Up CRON Jobs

### Basic Setup

1. **Make script executable:**
   ```bash
   chmod +x /home/ubuntu/zoom_attendance_cron.py
   ```

2. **Edit crontab:**
   ```bash
   crontab -e
   ```

3. **Add your schedule:**
   ```bash
   # Tuesday at 10:00 PM PST
   0 22 * * 2 /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py
   
   # Saturday at 1:30 PM PST
   30 13 * * 6 /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py
   ```

### Common CRON Schedules

```bash
# Daily at 9 AM
0 9 * * * /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py

# Every weekday at 5 PM
0 17 * * 1-5 /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py

# Every 6 hours
0 */6 * * * /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py

# First Monday of every month at 9 AM
0 9 1-7 * 1 /usr/bin/python3 /home/ubuntu/zoom_attendance_cron.py
```

### Timezone Setup (Recommended)

Set your server to PST timezone:
```bash
sudo timedatectl set-timezone America/Los_Angeles
timedatectl  # Verify the change
```

## 📊 Output Examples

### Excel Report
| Meeting Date (PST) | Meeting Time (PST) | Participant Name |
|-------------------|-------------------|------------------|
| 2025-06-12        | 10:00            | Alice Johnson    |
| 2025-06-12        | 10:00            | Bob Smith        |
| 2025-06-12        | 10:00            | Carol Davis      |

### Email Body
```
Hello,

Please find attached the Zoom attendance report.

Generated: 2025-06-12 at 14:30 PST
Meeting ID: 123456789
Time Range: 24h

Participants:
Alice Johnson
Bob Smith
Carol Davis
David Wilson

Best regards,
Zoom Attendance System
```

## 🔍 Monitoring & Troubleshooting

### View CRON Jobs
```bash
crontab -l
```

### Check CRON Logs
```bash
sudo tail -f /var/log/syslog | grep CRON
```

### Check Script Logs
```bash
tail -f /tmp/zoom_attendance.log
```

### Test Email Configuration
Use the interactive script's test email feature:
```bash
python3 zoom.py
# Select option 2: Test email
```

### Common Issues

**"Missing Zoom credentials" error:**
- Verify your `.env` file exists and has correct values
- Check that environment variables are properly set

**"Authentication failed" error:**
- Verify your Zoom API credentials
- Ensure your Server-to-Server OAuth app is activated

**Email fails:**
- For Gmail, use App Passwords instead of your regular password
- Check SMTP server settings
- Verify recipient email addresses

**No meetings found:**
- Check the time range (meetings must be within the specified time)
- Verify the meeting ID is correct (use quotes for strings)
- Ensure meetings actually occurred within the time window

## 🔒 Security Best Practices

1. **Protect your .env file:**
   ```bash
   chmod 600 .env
   ```

2. **Use App Passwords for Gmail** instead of your regular password

3. **Regularly rotate API credentials** and passwords

4. **Don't commit .env files** to version control - add to `.gitignore`

## 📁 File Structure

```
zoom-attendance/
├── zoom.py                    # Interactive version
├── zoom_attendance_cron.py    # CRON version
├── .env                       # Environment variables (create this)
├── .env.template             # Template for environment variables
└── README.md                 # This file
```

## 🤝 Support

If you encounter issues:

1. **Check the logs** first (`/tmp/zoom_attendance.log`)
2. **Test manually** before setting up CRON
3. **Verify all credentials** and configuration
4. **Check network connectivity** and firewall settings

## 📄 License

This project is open source. Feel free to modify and adapt for your needs.

---

*Generated attendance reports help track meeting participation and ensure accountability in remote work environments.*# ZoomAttendanceBot
Zoom Attendance Emailer




