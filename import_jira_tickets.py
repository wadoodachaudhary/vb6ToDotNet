#!/usr/bin/env python3
import os
import sys
import json
import urllib.request
import urllib.parse
import base64
import uuid
import re
import glob
from datetime import datetime

def sanitize_filename(name):
    """Sanitize the username to be safe in file names, mirroring C# behaviour."""
    cleaned = ''.join(c if c.isalnum() or c in (' ', '_', '-') else '-' for c in name.strip())
    cleaned = re.sub(r'[\s_-]+', '-', cleaned)
    cleaned = cleaned.strip('-')
    if len(cleaned) > 40:
        cleaned = cleaned[:40]
    return cleaned if cleaned else 'user'

def parse_jira_date(date_str):
    """Parse Jira's ISO-8601 date string with timezone offset and convert to C#-compatible format."""
    # Jira date format: e.g. "2026-06-24T16:14:30.000-0400"
    # We clean the timezone offset for python datetime parsing or keep as string
    try:
        # Try parsing with datetime.fromisoformat if python version >= 3.7
        # Normalizing tz offset format from -0400 to -04:00
        if re.search(r'[\+\-]\d{4}$', date_str):
            date_str = date_str[:-2] + ':' + date_str[-2:]
        return datetime.fromisoformat(date_str)
    except Exception:
        # Fallback to simple parse or current time if it fails
        try:
            clean_str = re.sub(r'\.\d+', '', date_str) # strip milliseconds
            clean_str = clean_str.split('+')[0].split('-')[0].strip() # strip tz
            return datetime.strptime(clean_str, "%Y-%m-%dT%H:%M:%S")
        except Exception:
            return datetime.now()

def compose_body(item):
    """Build the body content of the bug report, mirroring C# ComposeBody."""
    parts = []
    if item.get('EnvironmentInfo'):
        parts.append(f"Environment:\n{item['EnvironmentInfo'].strip()}")
    if item.get('Description'):
        parts.append(f"Description:\n{item['Description'].strip()}")
    if item.get('Prerequisites'):
        parts.append(f"Prerequisites:\n{item['Prerequisites'].strip()}")
    if item.get('StepsToReproduce'):
        parts.append(f"Steps to Reproduce:\n{item['StepsToReproduce'].strip()}")
    if item.get('ActualResult'):
        parts.append(f"Actual Result:\n{item['ActualResult'].strip()}")
    if item.get('ExpectedResult'):
        parts.append(f"Expected Result:\n{item['ExpectedResult'].strip()}")
    
    mig = []
    if item.get('VB6Form'):
        mig.append(f"VB6 Form: {item['VB6Form']}")
    if item.get('Module'):
        mig.append(f"Module: {item['Module']}")
    if mig:
        parts.append("Migration context: " + "  |  ".join(mig))
    return "\n\n".join(parts)

def build_printable_html(item, base64_screenshot):
    """Build the printable HTML document format matching FeedbackService.cs."""
    def e(s):
        if not s: return ""
        return (str(s).replace('&', '&amp;')
                      .replace('<', '&lt;')
                      .replace('>', '&gt;')
                      .replace('"', '&quot;')
                      .replace("'", '&#x27;'))

    img_tag = ("<p><em>(no screenshot)</em></p>" if not base64_screenshot 
               else f'<img src="{base64_screenshot}" style="max-width:100%;border:1px solid #999;" />')

    due_date_str = item.get('DueDate')[:10] if item.get('DueDate') else ""

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8" />
<title>Ticket #{item['TicketNumber']} — {e(item['Topic'])}</title>
<style>
 body {{ font-family: 'Segoe UI', Tahoma, sans-serif; color:#222; margin:24px; }}
 h1 {{ font-size:18px; border-bottom:2px solid #316ac5; padding-bottom:6px; }}
 table.meta {{ border-collapse:collapse; margin:12px 0; }}
 table.meta td {{ padding:4px 10px; vertical-align:top; }}
 table.meta td.k {{ font-weight:bold; color:#555; white-space:nowrap; }}
 .desc {{ white-space:pre-wrap; background:#f6f6f6; border:1px solid #ddd; padding:10px; }}
</style></head>
<body>
 <h1>Ticket #{item['TicketNumber']} — {e(item['Topic'])}</h1>
 <table class="meta">
  <tr><td class="k">Submitted By</td><td>{e(item['UserName'])}</td></tr>
  <tr><td class="k">Date / Time</td><td>{item['CreatedDate']}</td></tr>
  <tr><td class="k">Issue type</td><td>{e(item['IssueType'])}</td></tr>
  <tr><td class="k">Priority / Severity</td><td>{e(item['Priority'])} / {e(item['Severity'])}</td></tr>
  <tr><td class="k">Status</td><td>{e(item['Status'])}</td></tr>
  <tr><td class="k">Labels</td><td>{e(item['Labels'])}</td></tr>
  <tr><td class="k">Category</td><td>{e(item['Category'])}</td></tr>
  <tr><td class="k">Due date</td><td>{due_date_str}</td></tr>
  <tr><td class="k">Fix version / Sprint</td><td>{e(item['FixVersion'])}</td></tr>
  <tr><td class="k">VB6 form / Module</td><td>{e(item['VB6Form'])} / {e(item['Module'])}</td></tr>
  <tr><td class="k">Migrated page</td><td>{e(item['ScreenName'])}</td></tr>
  <tr><td class="k">Focused control</td><td>{e(item['FocusedControl'])}</td></tr>
  <tr><td class="k">Origin URL</td><td>{e(item['Url'])}</td></tr>
 </table>
 <h3>Bug report</h3>
 <div class="desc">{e(compose_body(item))}</div>
 <h3>Annotated Screenshot</h3>
 {img_tag}
</body></html>"""

def build_llm_markdown(item):
    """Build the Markdown companion file matching FeedbackService.cs."""
    lines = [
        f"# LLM Debugging Instructions: Ticket #{item['TicketNumber']:05d}",
        "",
        f"**Topic:** {item['Topic']}",
        f"**Author:** {item['UserName']}",
        f"**Timestamp:** {item['CreatedDate']}",
        f"**Active URL:** {item['Url']}",
        f"**Screen/Form:** `{item['ScreenName']}`",
        f"**Focused Control:** `{item['FocusedControl']}`",
        "",
        "---",
        "",
        "## 1. Problem Description",
        item['Description'] or "No description provided.",
        "",
        "---",
        "",
        "## 2. Screenshot & Visual Context"
    ]
    if item.get('ScreenshotPath'):
        lines.append(f"![Screenshot Companion]({item['ScreenshotPath']})")
    else:
        lines.append("*(No screenshot companion attached)*")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 3. Environment Context")
    lines.append("*(No diagnostics logs gathered)*")
    lines.append("")
    return "\n".join(lines)

def make_jira_request(url, headers):
    """Make HTTP request using urllib standard library."""
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as response:
            return response.read()
    except urllib.error.HTTPError as e:
        print(f"HTTP Error: {e.code} - {e.reason}")
        try:
            print(e.read().decode('utf-8'))
        except Exception:
            pass
        raise
    except Exception as e:
        print(f"Network error calling Jira: {e}")
        raise

def main():
    print("====================================================")
    print(" Jira to HomeFront Feedback Tickets Import Utility ")
    print("====================================================\n")

    # Path setup relative to script location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    tickets_dir = os.path.join(script_dir, "HomeFrontPB", "wwwroot", "tickets")

    if not os.path.exists(tickets_dir):
        # Let's search if HomeFrontPB is situated differently
        possible_dir = os.path.join(script_dir, "wwwroot", "tickets")
        if os.path.exists(possible_dir):
            tickets_dir = possible_dir
        else:
            # Let's fallback to current workspace's tickets folder
            tickets_dir = "/Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontPB/wwwroot/tickets"

    print(f"Target Tickets Directory: {tickets_dir}")
    if not os.path.exists(tickets_dir):
        try:
            os.makedirs(tickets_dir, exist_ok=True)
            print("Created tickets directory.")
        except Exception as e:
            print(f"Error creating directory: {e}")
            sys.exit(1)

    # 1. Load credentials
    # Try reading from config file App_Data/jira-settings.json first if exists
    app_data_settings = os.path.join(script_dir, "HomeFrontPB", "App_Data", "jira-settings.json")
    jira_url = "https://innovatixinc.atlassian.net"
    email = ""
    api_token = ""
    project_key = "HHM"

    if os.path.exists(app_data_settings):
        try:
            with open(app_data_settings, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                jira_url = settings.get("BaseUrl", jira_url).rstrip('/')
                email = settings.get("Email", email)
                api_token = settings.get("ApiToken", api_token)
                project_key = settings.get("ProjectKey", project_key)
                print(f"Loaded configuration from {app_data_settings}")
        except Exception as e:
            print(f"Warning: Failed to load config from {app_data_settings}: {e}")

    # Prompt for missing credentials if not loaded
    if not email:
        email = input("Enter Jira Email address: ").strip()
    if not api_token:
        # Prompt securely or normal input since this is a terminal command run by the user
        api_token = input("Enter Jira API Token (generate at id.atlassian.net): ").strip()

    if not email or not api_token:
        print("Error: Email and API Token are required.")
        sys.exit(1)

    # Setup headers
    auth_str = f"{email}:{api_token}"
    auth_b64 = base64.b64encode(auth_str.encode('utf-8')).decode('utf-8')
    headers = {
        'Authorization': f'Basic {auth_b64}',
        'Accept': 'application/json',
        'User-Agent': 'HomeFront Jira Sync'
    }

    # 2. Get list of already imported Jira issues to prevent duplicates
    print("\nScanning existing HomeFront tickets for duplicates...")
    imported_keys = set()
    max_ticket_no = 0

    json_files = glob.glob(os.path.join(tickets_dir, "ticket_*.json"))
    for filepath in json_files:
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                ticket = json.load(f)
                t_num = ticket.get('TicketNumber', 0)
                if t_num > max_ticket_no:
                    max_ticket_no = t_num
                
                url = ticket.get('Url', '')
                # Find Jira Issue key like HHM-123 in the url
                match = re.search(r'/browse/([A-Z0-9]+-\d+)', url)
                if match:
                    imported_keys.add(match.group(1))
        except Exception as e:
            print(f"  Skipping parse error in {os.path.basename(filepath)}: {e}")

    print(f"Found {len(imported_keys)} already imported Jira tickets.")
    print(f"Current maximum ticket number: #{max_ticket_no}")

    # 3. Query Jira Cloud
    # Fetch issues for HHM project sorted by creation date so new ticket numbers follow chronological order
    jql = f"project = {project_key} ORDER BY created ASC"
    print(f"\nQuerying Jira issues with JQL: '{jql}'...")
    
    encoded_jql = urllib.parse.quote(jql)
    # Using REST API v2 as it returns plain-text descriptions instead of complex v3 ADF structures
    url = f"{jira_url}/rest/api/2/search?jql={encoded_jql}&maxResults=100&fields=summary,description,created,duedate,status,priority,issuetype,reporter,assignee,labels,components,fixVersions,attachment"

    try:
        resp_bytes = make_jira_request(url, headers)
        search_res = json.loads(resp_bytes.decode('utf-8'))
    except Exception as e:
        print(f"Failed to query Jira: {e}")
        sys.exit(1)

    issues = search_res.get('issues', [])
    print(f"Jira search returned {len(issues)} issues.")

    newly_imported = 0
    for issue in issues:
        key = issue['key']
        if key in imported_keys:
            # Already imported
            continue

        print(f"\nProcessing issue {key}...")
        
        fields = issue.get('fields', {})
        summary = fields.get('summary', '')
        raw_description = fields.get('description', '') or 'No description.'
        
        # Parse Dates
        created_raw = fields.get('created', '')
        created_dt = parse_jira_date(created_raw)
        created_str = created_dt.strftime("%Y-%m-%dT%H:%M:%S.%f-04:00") # default to -04:00 matching user timezone
        
        due_raw = fields.get('duedate')
        due_str = f"{due_raw}T00:00:00-04:00" if due_raw else None

        # Map Status
        jira_status = fields.get('status', {}).get('name', 'To Do')
        status_map = {
            "backlog": "To Do",
            "to do": "To Do",
            "open": "To Do",
            "new": "To Do",
            "selected for development": "To Do",
            "in progress": "In Progress",
            "active": "In Progress",
            "under review": "In Review",
            "in review": "In Review",
            "qa": "In Review",
            "resolved": "In Review",
            "done": "Done",
            "closed": "Done",
            "complete": "Done",
        }
        status = status_map.get(jira_status.lower(), jira_status)

        # Map Priority
        jira_priority = fields.get('priority', {}).get('name', 'Medium')
        priority_map = {
            "highest": "High",
            "high": "High",
            "medium": "Medium",
            "low": "Low",
            "lowest": "Low",
        }
        priority = priority_map.get(jira_priority.lower(), jira_priority)

        # Map Categories/Components
        components = fields.get('components', [])
        category = components[0].get('name', '') if components else ''

        # Map Assignee / Reporter
        reporter_name = fields.get('reporter', {}).get('displayName', 'Jira Reporter')
        assignee_name = fields.get('assignee', {}).get('displayName', '') or 'Unassigned'

        # Map Labels
        labels_list = fields.get('labels', [])
        labels = ", ".join(labels_list) if labels_list else ""

        # Map Fix Version
        fix_versions = fields.get('fixVersions', [])
        fix_version = fix_versions[0].get('name', '') if fix_versions else ''

        # Map Issue Type
        issue_type = fields.get('issuetype', {}).get('name', 'Bug')

        # Increment ticket number
        max_ticket_no += 1
        ticket_number = max_ticket_no

        # Handle screenshots/attachments
        attachments = fields.get('attachment', [])
        image_attachment = None
        for att in attachments:
            mime = att.get('mimeType', '')
            if mime.startswith('image/'):
                image_attachment = att
                break

        screenshot_path = ""
        base64_screenshot = ""

        safe_reporter = sanitize_filename(reporter_name)
        stamp = created_dt.strftime("%Y%m%d-%H%M%S")
        base_name = f"ticket_{ticket_number:05d}_{safe_reporter}_{stamp}"

        if image_attachment:
            att_url = image_attachment['content']
            att_filename = f"{base_name}.png"
            att_filepath = os.path.join(tickets_dir, att_filename)
            
            print(f"  Downloading screenshot attachment: '{image_attachment['filename']}'...")
            try:
                img_data = make_jira_request(att_url, headers)
                with open(att_filepath, 'wb') as img_f:
                    img_f.write(img_data)
                
                screenshot_path = f"/tickets/{att_filename}"
                b64_encoded = base64.b64encode(img_data).decode('utf-8')
                base64_screenshot = f"data:{image_attachment['mimeType']};base64,{b64_encoded}"
                print("  Screenshot saved successfully.")
            except Exception as e:
                print(f"  Warning: failed to download screenshot attachment: {e}")

        # Construct FeedbackItem dict
        item = {
            "TicketNumber": ticket_number,
            "Id": uuid.uuid4().hex,
            "Topic": f"[{key}] {summary}",
            "Description": raw_description,
            "UserName": reporter_name,
            "CreatedDate": created_str,
            "Url": f"{jira_url}/browse/{key}",
            "Status": status,
            "Priority": priority,
            "Category": category,
            "Assignee": assignee_name,
            "IssueType": issue_type,
            "Severity": "Major",
            "Labels": labels,
            "FixVersion": fix_version,
            "DueDate": due_str,
            "EnvironmentInfo": "",
            "Prerequisites": "",
            "StepsToReproduce": "",
            "ActualResult": "",
            "ExpectedResult": "",
            "VB6Form": "",
            "Module": "",
            "ScreenName": "",
            "FocusedControl": "",
            "ScreenshotPath": screenshot_path,
            "DocumentPath": f"/tickets/{base_name}.html",
            "AnnotationsJson": "",
            "DiagnosticsJson": ""
        }

        # Write JSON file
        json_path = os.path.join(tickets_dir, f"{base_name}.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(item, f, indent=2)

        # Write HTML file
        html_path = os.path.join(tickets_dir, f"{base_name}.html")
        html_content = build_printable_html(item, base64_screenshot)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Write Markdown file
        md_path = os.path.join(tickets_dir, f"{base_name}.md")
        md_content = build_llm_markdown(item)
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

        print(f"  Imported as HomeFront Ticket #{ticket_number:05d}")
        newly_imported += 1

    print("\n====================================================")
    print(f" Done! Imported {newly_imported} new tickets from Jira.")
    print("====================================================")

if __name__ == "__main__":
    main()
