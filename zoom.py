#!/usr/bin/env python3
"""
Zoom Attendance Report Generator

Generate attendance reports from Zoom meetings with PST timezone support.
Supports time ranges in hours/minutes and deduplicates multiple sessions.
"""

import os
import json
import smtplib
import base64
import requests
import pandas as pd
import urllib.parse
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Dict, Optional
import pytz

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    if os.path.exists('.env'):
        with open('.env', 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip().strip('"').strip("'")


class TimeHelper:
    """Timezone and time parsing utilities"""
    
    @staticmethod
    def parse_time_input(time_input: str) -> timedelta:
        """Parse '2h', '30m', '3d' into timedelta"""
        time_input = time_input.strip().lower()
        if not time_input:
            return timedelta(days=1)
        
        if time_input[-1] in ['h', 'm', 'd']:
            try:
                number = int(time_input[:-1])
                unit = time_input[-1]
                if unit == 'h':
                    return timedelta(hours=number)
                elif unit == 'm':
                    return timedelta(minutes=number)
                elif unit == 'd':
                    return timedelta(days=number)
            except ValueError:
                pass
        
        try:
            return timedelta(days=int(time_input))
        except ValueError:
            return timedelta(days=1)
    
    @staticmethod
    def to_pst(utc_time_str: str) -> str:
        """Convert UTC time to PST"""
        if not utc_time_str or utc_time_str == "Unknown":
            return "Unknown"
        
        try:
            utc_dt = datetime.strptime(utc_time_str[:19], "%Y-%m-%dT%H:%M:%S")
            utc_dt = pytz.utc.localize(utc_dt)
            pst_dt = utc_dt.astimezone(pytz.timezone('America/Los_Angeles'))
            return pst_dt.strftime("%Y-%m-%d %H:%M PST")
        except:
            return utc_time_str
    
    @staticmethod
    def to_pst_time_only(utc_time_str: str) -> str:
        """Convert UTC time to PST time only (HH:MM)"""
        if not utc_time_str or utc_time_str == "Unknown":
            return "Unknown"
        
        try:
            utc_dt = datetime.strptime(utc_time_str[:19], "%Y-%m-%dT%H:%M:%S")
            utc_dt = pytz.utc.localize(utc_dt)
            pst_dt = utc_dt.astimezone(pytz.timezone('America/Los_Angeles'))
            return pst_dt.strftime("%H:%M")
        except:
            return utc_time_str[:5] if len(utc_time_str) > 5 else utc_time_str


class ZoomAPI:
    """Zoom API client"""
    
    def __init__(self, account_id: str, client_id: str, client_secret: str):
        self.account_id = account_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.access_token = None
        self.base_url = "https://api.zoom.us/v2"
    
    def authenticate(self):
        """Get OAuth access token"""
        url = "https://zoom.us/oauth/token"
        credentials = base64.b64encode(f"{self.client_id}:{self.client_secret}".encode()).decode()
        
        response = requests.post(url, 
            headers={
                "Authorization": f"Basic {credentials}",
                "Content-Type": "application/x-www-form-urlencoded"
            },
            data={
                "grant_type": "account_credentials",
                "account_id": self.account_id
            }
        )
        
        if response.status_code == 200:
            self.access_token = response.json()["access_token"]
        else:
            raise Exception(f"Authentication failed: {response.text}")
    
    def get_headers(self):
        """Get authorization headers"""
        if not self.access_token:
            self.authenticate()
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
    
    def get_meetings(self, time_delta: timedelta) -> List[Dict]:
        """Get meetings within time range"""
        hours = time_delta.total_seconds() / 3600
        print(f"Getting meetings from last {hours:.1f} hours...")
        
        response = requests.get(
            f"{self.base_url}/users/me/meetings",
            headers=self.get_headers(),
            params={"type": "previous_meetings", "page_size": 300}
        )
        
        if response.status_code != 200:
            print(f"Failed to get meetings: {response.text}")
            return []
        
        all_meetings = response.json().get("meetings", [])
        cutoff_time = datetime.now() - time_delta
        filtered_meetings = []
        
        for meeting in all_meetings:
            start_time = meeting.get("start_time", "")
            if not start_time:
                continue
            
            try:
                meeting_dt = datetime.strptime(start_time[:19], "%Y-%m-%dT%H:%M:%S")
                if meeting_dt >= cutoff_time:
                    filtered_meetings.append(meeting)
            except ValueError:
                continue
        
        print(f"Found {len(filtered_meetings)} meetings within time range")
        return filtered_meetings
    
    def get_meeting_instances(self, meeting_id: str, time_delta: timedelta) -> List[Dict]:
        """Get instances of a specific meeting within time range"""
        instances = []
        
        # Try past meetings endpoint
        response = requests.get(
            f"{self.base_url}/past_meetings/{meeting_id}/instances",
            headers=self.get_headers()
        )
        
        if response.status_code == 200:
            instances = response.json().get("meetings", [])
        
        # Also check user meetings
        user_meetings = self.get_meetings(time_delta)
        for meeting in user_meetings:
            if str(meeting.get("id")) == str(meeting_id):
                instances.append(meeting)
        
        # Filter by time and deduplicate
        cutoff_time = datetime.now() - time_delta
        unique_instances = {}
        
        for instance in instances:
            start_time = instance.get("start_time", "")
            if not start_time:
                continue
            
            try:
                meeting_dt = datetime.strptime(start_time[:19], "%Y-%m-%dT%H:%M:%S")
                if meeting_dt >= cutoff_time:
                    uuid = instance.get("uuid", instance.get("id"))
                    if uuid:
                        unique_instances[uuid] = instance
            except ValueError:
                continue
        
        return list(unique_instances.values())
    
    def get_participants(self, meeting_uuid: str) -> List[Dict]:
        """Get participants for a meeting"""
        encoded_uuid = urllib.parse.quote(meeting_uuid, safe='')
        response = requests.get(
            f"{self.base_url}/report/meetings/{encoded_uuid}/participants",
            headers=self.get_headers(),
            params={"page_size": 300}
        )
        
        if response.status_code == 200:
            return response.json().get("participants", [])
        return []


class AttendanceReporter:
    """Generate attendance reports"""
    
    def __init__(self, zoom_api: ZoomAPI):
        self.zoom_api = zoom_api
    
    def generate_report(self, meeting_id: str = None, time_delta: timedelta = timedelta(days=1)) -> tuple[pd.DataFrame, list]:
        """Generate attendance report - returns DataFrame and participant list"""
        if meeting_id:
            meetings = self.zoom_api.get_meeting_instances(meeting_id, time_delta)
            print(f"Found {len(meetings)} instances of meeting {meeting_id}")
        else:
            meetings = self.zoom_api.get_meetings(time_delta)
            print(f"Found {len(meetings)} meetings")
        
        if not meetings:
            return pd.DataFrame(), []
        
        attendance_data = []
        all_participants = []
        
        for i, meeting in enumerate(meetings, 1):
            topic = meeting.get("topic", "Unknown Meeting")
            start_time = meeting.get("start_time", "")
            meeting_uuid = meeting.get("uuid", meeting.get("id"))
            
            print(f"Processing {i}/{len(meetings)}: {topic}")
            
            participants = self.zoom_api.get_participants(str(meeting_uuid))
            
            if not participants:
                attendance_data.append({
                    "Meeting Date (PST)": TimeHelper.to_pst(start_time)[:10] if start_time else "Unknown",
                    "Meeting Time (PST)": TimeHelper.to_pst_time_only(start_time),
                    "Participant Name": "No attendees"
                })
                all_participants.append("No attendees")
            else:
                deduplicated = self._deduplicate_participants(participants)
                
                for participant in deduplicated:
                    participant_name = participant.get("name", "Unknown")
                    attendance_data.append({
                        "Meeting Date (PST)": TimeHelper.to_pst(start_time)[:10] if start_time else "Unknown",
                        "Meeting Time (PST)": TimeHelper.to_pst_time_only(start_time),
                        "Participant Name": participant_name
                    })
                    all_participants.append(participant_name)
        
        print(f"Generated {len(attendance_data)} attendance records")
        return pd.DataFrame(attendance_data), all_participants
    
    def _deduplicate_participants(self, participants: List[Dict]) -> List[Dict]:
        """Combine multiple sessions for same participant"""
        participant_groups = {}
        
        for participant in participants:
            email = participant.get("user_email", "").strip().lower()
            name = participant.get("name", "").strip()
            key = email if email and email != "n/a" else name.lower()
            
            if key not in participant_groups:
                participant_groups[key] = []
            participant_groups[key].append(participant)
        
        deduplicated = []
        for sessions in participant_groups.values():
            if len(sessions) == 1:
                deduplicated.append(sessions[0])
            else:
                combined = self._combine_sessions(sessions)
                deduplicated.append(combined)
                print(f"  Combined {len(sessions)} sessions for {combined.get('name', 'Unknown')}")
        
        return deduplicated
    
    def _combine_sessions(self, sessions: List[Dict]) -> Dict:
        """Combine multiple sessions for same participant"""
        sessions.sort(key=lambda x: x.get("join_time", ""))
        
        total_duration = sum(session.get("duration", 0) for session in sessions)
        best_email = next((s.get("user_email", "") for s in sessions 
                          if s.get("user_email", "").strip() and s.get("user_email", "").lower() != "n/a"), "N/A")
        
        return {
            "name": sessions[0].get("name", "Unknown"),
            "user_email": best_email,
            "join_time": sessions[0].get("join_time", ""),
            "leave_time": sessions[-1].get("leave_time", ""),
            "duration": total_duration,
            "status": sessions[0].get("status", "Unknown")
        }
    
    def save_report(self, df: pd.DataFrame) -> str:
        """Save report to Excel"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"zoom_attendance_report_{timestamp}.xlsx"
        
        if len(df) == 0:
            df = pd.DataFrame({"Message": ["No data found"]})
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Attendance Report', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Attendance Report']
            for column in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in column)
                worksheet.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)
        
        print(f"Report saved: {filename}")
        return filename


class EmailSender:
    """Email functionality"""
    
    def __init__(self, smtp_server: str, smtp_port: int, email: str, password: str):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email = email
        self.password = password
    
    def send_report(self, filename: str, recipients: List[str], participants: List[str]):
        """Send report via email with participant list in body"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.email
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = f"Zoom Attendance Report - {datetime.now().strftime('%Y-%m-%d')}"
            
            # Create participant list for email body
            unique_participants = list(set(participants))  # Remove duplicates
            unique_participants.sort()  # Sort alphabetically
            participant_list = '\n'.join(unique_participants) if unique_participants else "No participants found"
            
            body = f"""Hello,

Please find attached the Zoom attendance report.

Generated: {datetime.now().strftime('%Y-%m-%d at %H:%M PST')}

Participants:
{participant_list}

Best regards,
Zoom Attendance System"""
            
            msg.attach(MIMEText(body, 'plain'))
            
            with open(filename, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filename)}')
            msg.attach(part)
            
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.email, self.password)
            server.sendmail(self.email, recipients, msg.as_string())
            server.quit()
            
            print(f"Email sent to: {', '.join(recipients)}")
            return True
            
        except Exception as e:
            print(f"Email failed: {e}")
            return False


def get_config():
    """Get configuration from environment"""
    def safe_int(value, default):
        try:
            return int(str(value).strip().strip('"').strip("'"))
        except:
            return default
    
    return {
        'zoom_account_id': os.getenv('ZOOM_ACCOUNT_ID'),
        'zoom_client_id': os.getenv('ZOOM_CLIENT_ID'),
        'zoom_client_secret': os.getenv('ZOOM_CLIENT_SECRET'),
        'sender_email': os.getenv('SENDER_EMAIL'),
        'sender_password': os.getenv('SENDER_PASSWORD'),
        'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
        'smtp_port': safe_int(os.getenv('SMTP_PORT', '587'), 587),
        'recipients': [x.strip() for x in os.getenv('EMAIL_RECIPIENTS', '').split(',') if x.strip()]
    }


def manual_report():
    """Generate manual report"""
    config = get_config()
    
    if not all([config['zoom_account_id'], config['zoom_client_id'], config['zoom_client_secret']]):
        print("Missing Zoom credentials. Set ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET")
        return
    
    meeting_id = input("Enter meeting ID (or press Enter for all meetings): ").strip() or None
    
    print("\nTime range examples: 2h, 30m, 1d")
    time_input = input("Time to look back (default 1d): ").strip() or "1d"
    time_delta = TimeHelper.parse_time_input(time_input)
    
    hours = time_delta.total_seconds() / 3600
    cutoff_pst = TimeHelper.to_pst((datetime.now() - time_delta).strftime("%Y-%m-%dT%H:%M:%S"))
    print(f"\nLooking for meetings from last {hours:.1f} hours (after {cutoff_pst})")
    
    try:
        zoom_api = ZoomAPI(config['zoom_account_id'], config['zoom_client_id'], config['zoom_client_secret'])
        reporter = AttendanceReporter(zoom_api)
        df, participants = reporter.generate_report(meeting_id, time_delta)
        filename = reporter.save_report(df)
        
        if input("\nSend via email? (y/n): ").lower() == 'y':
            if not config['sender_email'] or not config['sender_password']:
                print("Missing email credentials. Set SENDER_EMAIL, SENDER_PASSWORD")
                return
            
            recipients = config['recipients']
            if not recipients:
                recipients_input = input("Enter recipient emails (comma-separated): ").strip()
                recipients = [x.strip() for x in recipients_input.split(',') if x.strip()]
            
            if recipients:
                email_sender = EmailSender(
                    config['smtp_server'], config['smtp_port'],
                    config['sender_email'], config['sender_password']
                )
                email_sender.send_report(filename, recipients, participants)
        
        print("Done!")
        
    except Exception as e:
        print(f"Error: {e}")
        if "pytz" in str(e):
            print("Install timezone library: pip install pytz")


def test_email():
    """Test email configuration"""
    config = get_config()
    
    email = config['sender_email'] or input("Email: ").strip()
    password = config['sender_password'] or input("Password: ").strip()
    recipient = input("Test recipient: ").strip()
    
    try:
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = recipient
        msg['Subject'] = "Test Email - Zoom Attendance System"
        msg.attach(MIMEText("Test email successful!", 'plain'))
        
        server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
        server.starttls()
        server.login(email, password)
        server.sendmail(email, [recipient], msg.as_string())
        server.quit()
        
        print("Test email sent successfully!")
        
    except Exception as e:
        print(f"Test failed: {e}")


def create_env_template():
    """Create .env template"""
    template = """# Zoom API Configuration
ZOOM_ACCOUNT_ID=your_account_id
ZOOM_CLIENT_ID=your_client_id
ZOOM_CLIENT_SECRET=your_client_secret

# Email Configuration
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_app_password
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Recipients (comma-separated)
EMAIL_RECIPIENTS=manager@company.com,hr@company.com
"""
    
    with open('.env.template', 'w') as f:
        f.write(template)
    print("Created .env.template - copy to .env and fill in your values")


def show_config():
    """Show configuration status"""
    config = get_config()
    
    print("Configuration Status:")
    print(f"Zoom Account ID: {'✓' if config['zoom_account_id'] else '✗'}")
    print(f"Zoom Client ID: {'✓' if config['zoom_client_id'] else '✗'}")
    print(f"Zoom Client Secret: {'✓' if config['zoom_client_secret'] else '✗'}")
    print(f"Sender Email: {config['sender_email'] or '✗ Not set'}")
    print(f"Recipients: {len(config['recipients'])} configured")


def main():
    """Main menu"""
    print("Zoom Attendance Report Generator")
    print("\n1. Generate report")
    print("2. Test email")
    print("3. Create .env template")
    print("4. Show configuration")
    
    choice = input("\nSelect (1-4): ").strip()
    
    if choice == "1":
        manual_report()
    elif choice == "2":
        test_email()
    elif choice == "3":
        create_env_template()
    elif choice == "4":
        show_config()
    else:
        print("Invalid choice")


if __name__ == "__main__":
    main()