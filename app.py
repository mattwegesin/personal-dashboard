import os
import sys
import json
import time
import requests
import subprocess
from flask import Flask, render_template, jsonify, request
from dotenv import load_dotenv
from google import genai
from google.genai import types

# Dynamically import system instructions from sibling project (network_design_web) to prevent duplication!
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'network_design_web')))
try:
    from app import NETWORK_DESIGNER_INSTRUCTION, READINESS_INSTRUCTION
except Exception as e:
    print(f"Warning: Could not dynamically import instructions from sibling project: {e}")
    # Robust fallback instructions
    NETWORK_DESIGNER_INSTRUCTION = "You are a professional Wi-Fi Network Designer. Create a Wi-Fi design, BOM, SOW, and Mermaid diagram."
    READINESS_INSTRUCTION = "You are a Wi-Fi Design Intake Specialist. Audit SharePoint folder files for floorplans, materials, and scales."

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-dashboard-key-12345')

# Helper: Find the absolute path of a file with sibling/upward directory fallback
def find_workspace_file(filename, subfolder=None):
    # Try direct sibling first
    current_dir = os.path.abspath(os.path.dirname(__file__))
    
    # 1. Direct sibling directory check (e.g. personal-dashboard/../subfolder/filename)
    if subfolder:
        path = os.path.abspath(os.path.join(current_dir, '..', subfolder, filename))
    else:
        path = os.path.abspath(os.path.join(current_dir, '..', filename))
        
    if os.path.exists(path):
        return path
        
    # 2. Parallel workspace check (e.g. searching from parent)
    parent_dir = os.path.dirname(current_dir)
    if subfolder:
        path = os.path.abspath(os.path.join(parent_dir, subfolder, filename))
    else:
        path = os.path.abspath(os.path.join(parent_dir, filename))
        
    if os.path.exists(path):
        return path
        
    # 3. Double-upward backup (just in case directory structure is nested)
    grandparent_dir = os.path.dirname(parent_dir)
    if subfolder:
        path = os.path.abspath(os.path.join(grandparent_dir, subfolder, filename))
    else:
        path = os.path.abspath(os.path.join(grandparent_dir, filename))
        
    if os.path.exists(path):
        return path
        
    # Default fallback to original assumed sibling path
    if subfolder:
        return os.path.abspath(os.path.join(current_dir, '..', subfolder, filename))
    return os.path.abspath(os.path.join(current_dir, '..', filename))

# Auto-discover credentials from workspace siblings
def discover_credentials():
    creds = {
        'CLIENT_ID': os.environ.get('CLIENT_ID'),
        'TENANT_ID': os.environ.get('TENANT_ID', 'common'),
        'MONDAY_API_TOKEN': os.environ.get('MONDAY_API_TOKEN'),
        'BOARD_ID': os.environ.get('BOARD_ID', '9609739665')
    }
    
    # 1. Try loading from outlook-mcp/.env
    outlook_env = find_workspace_file('.env', 'outlook-mcp')
    if os.path.exists(outlook_env):
        try:
            with open(outlook_env, 'r') as f:
                for line in f:
                    if '=' in line and not line.strip().startswith('#'):
                        parts = line.strip().split('=', 1)
                        if len(parts) == 2:
                            k, v = parts[0].strip(), parts[1].strip().strip('"\'')
                            if k in ['CLIENT_ID', 'TENANT_ID'] and not creds[k]:
                                creds[k] = v
        except Exception as e:
            print(f"Error loading outlook-mcp/.env from {outlook_env}: {e}")
            
    # 2. Try loading Monday API token from monitor_monday.sh
    monitor_sh = find_workspace_file('monitor_monday.sh')
    if os.path.exists(monitor_sh) and not creds['MONDAY_API_TOKEN']:
        try:
            with open(monitor_sh, 'r') as f:
                for line in f:
                    if 'API_TOKEN=' in line:
                        parts = line.split('API_TOKEN=', 1)
                        if len(parts) == 2:
                            token = parts[1].strip().strip('"\';')
                            # Ensure we don't pick up comments or other bash logic
                            if token and not token.startswith('#') and len(token) > 20:
                                creds['MONDAY_API_TOKEN'] = token
                                break
        except Exception as e:
            print(f"Error loading Monday token from monitor_monday.sh: {e}")
            
    return creds

# Load environment variables
load_dotenv()
creds = discover_credentials()

# Helper: Get Outlook Token from outlook-mcp cache or refresh it
def get_outlook_token():
    client_id = creds.get('CLIENT_ID')
    tenant_id = creds.get('TENANT_ID', 'common')
    
    if not client_id:
        return None, "CLIENT_ID is missing"
        
    cache = None
    
    # 1. Try loading from environment variable (ideal for Render cloud hosting)
    env_cache = os.environ.get('OUTLOOK_TOKEN_CACHE_JSON')
    if env_cache:
        try:
            cache = json.loads(env_cache)
        except Exception as e:
            return None, f"Error parsing OUTLOOK_TOKEN_CACHE_JSON environment variable: {str(e)}"
            
    # 2. If not in env, fallback to loading the physical file (ideal for local development)
    if not cache:
        cache_path = find_workspace_file('token_cache.json', 'outlook-mcp')
        if not os.path.exists(cache_path):
            return None, f"MSAL token cache file not found. For local testing, ensure outlook-mcp/token_cache.json exists. For Render deployment, configure the OUTLOOK_TOKEN_CACHE_JSON environment variable with your local token cache contents. Searched at: {cache_path}"
            
        try:
            with open(cache_path, 'r') as f:
                cache = json.load(f)
        except Exception as e:
            return None, f"Error parsing physical MSAL cache: {str(e)}"
            
    try:
        # 1. Search for a valid, non-expired Access Token in cache
        access_tokens = cache.get('AccessToken', {})
        now_epoch = int(time.time())
        for key, token_obj in access_tokens.items():
            expires_on = int(token_obj.get('expires_on', 0))
            # 5-minute safety buffer
            if expires_on > now_epoch + 300:
                return token_obj.get('secret'), None
                
        # 2. If access token is expired, use the Refresh Token to get a new one
        refresh_tokens = cache.get('RefreshToken', {})
        for key, rt_obj in refresh_tokens.items():
            refresh_token = rt_obj.get('secret')
            if refresh_token:
                # Refresh token request
                url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
                data = {
                    'client_id': client_id,
                    'grant_type': 'refresh_token',
                    'refresh_token': refresh_token,
                    'scope': 'Mail.Read Mail.Send Calendars.Read Calendars.ReadWrite offline_access'
                }
                resp = requests.post(url, data=data)
                if resp.status_code == 200:
                    res_json = resp.json()
                    new_access_token = res_json.get('access_token')
                    # Optionally write back to token cache, but reading is sufficient for runtime use
                    return new_access_token, None
                else:
                    return None, f"Failed to refresh Outlook token: HTTP {resp.status_code} - {resp.text}"
                    
    except Exception as e:
        return None, f"Error parsing MSAL cache: {str(e)}"
        
    return None, "No active tokens found. Please run authentication in outlook-mcp."

# Route: Dashboard Homepage
@app.route('/')
def index():
    return render_template('index.html')

# API Route: Fetch Monday.com Projects
@app.route('/api/monday-projects')
def get_monday_projects():
    token = creds.get('MONDAY_API_TOKEN')
    board_id = creds.get('BOARD_ID')
    
    if not token:
        return jsonify({'error': 'Monday.com API Token is not configured.'}), 400
        
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": token,
        "Content-Type": "application/json"
    }
    
    # GraphQL query targeting specific groups, key column values, and subitems for design readiness
    query = """
    query {
      boards(ids: [%s]) {
        groups(ids: ["topics", "group_title", "group_mktk1nc2"]) {
          id
          title
          items_page(limit: 30) {
            items {
              id
              name
              column_values(ids: ["person", "color_mkvmrn92", "date4", "status", "link_mktk4wh4"]) {
                id
                text
                value
              }
              subitems {
                id
                name
                column_values(ids: ["status"]) {
                  id
                  text
                }
              }
            }
          }
        }
      }
    }
    """ % board_id
    
    try:
        resp = requests.post(url, json={"query": query}, headers=headers)
        if resp.status_code == 200:
            data = resp.json()
            if 'errors' in data:
                return jsonify({'error': data['errors'][0].get('message', 'GraphQL Error')}), 400
                
            boards = data.get('data', {}).get('boards', [])
            if not boards:
                return jsonify({'groups': []})
                
            groups = boards[0].get('groups', [])
            formatted_groups = []
            
            for g in groups:
                group_data = {
                    'id': g.get('id'),
                    'title': g.get('title'),
                    'projects': []
                }
                
                items = g.get('items_page', {}).get('items', [])
                for item in items:
                    proj = {
                        'id': item.get('id'),
                        'name': item.get('name'),
                        'sales_owner': 'N/A',
                        'flag': 'N/A',
                        'due_date': 'N/A',
                        'status': 'N/A',
                        'sp_link': '',
                        'readiness': 'NOT STARTED'  # Default red
                    }
                    
                    # 1. Parse standard column values
                    for cv in item.get('column_values', []):
                        col_id = cv.get('id')
                        text_val = cv.get('text') or 'N/A'
                        if col_id == 'person':
                            proj['sales_owner'] = text_val
                        elif col_id == 'color_mkvmrn92':
                            proj['flag'] = text_val
                        elif col_id == 'date4':
                            proj['due_date'] = text_val
                        elif col_id == 'status':
                            proj['status'] = text_val
                        elif col_id == 'link_mktk4wh4':
                            try:
                                val_json = json.loads(cv.get('value') or '{}')
                                proj['sp_link'] = val_json.get('url', '')
                            except:
                                proj['sp_link'] = ''
                            
                    # 2. Parse subitems to evaluate the project's exact workflow stage and readiness
                    subitems = item.get('subitems', []) or []
                    subitem_status = {}
                    for s in subitems:
                        s_name = s.get('name', '').strip()
                        status_val = 'Not Started'
                        for cv in s.get('column_values', []) or []:
                            if cv.get('id') == 'status':
                                status_val = cv.get('text') or 'Not Started'
                                break
                        subitem_status[s_name] = status_val
                    
                    # Evaluate based on the hierarchy of the engineering workflow
                    # A. Submission Stage (Update Salesforce)
                    salesforce_status = subitem_status.get('Update Salesforce') or subitem_status.get('Update SalesForce')
                    if salesforce_status:
                        if salesforce_status == 'Submitted for Approval':
                            proj['readiness'] = 'SUBMITTED'
                        elif salesforce_status == 'Done':
                            proj['readiness'] = 'READY'
                        elif salesforce_status in ['Working on it', 'Stuck', 'Waiting on Info']:
                            proj['readiness'] = 'MISSING INFO'
                        else:
                            proj['readiness'] = 'NOT STARTED'
                    
                    # B. Design/Quote Stage (Create design documents & quote)
                    elif 'Create design documents' in subitem_status or 'Create quote' in subitem_status:
                        docs_status = subitem_status.get('Create design documents', 'Not Started')
                        quote_status = subitem_status.get('Create quote', 'Not Started')
                        if docs_status == 'Done' or quote_status == 'Done':
                            proj['readiness'] = 'READY'
                        elif docs_status in ['Working on it', 'Waiting on Info', 'Stuck'] or quote_status in ['Working on it', 'Waiting on Info', 'Stuck']:
                            proj['readiness'] = 'MISSING INFO'
                        else:
                            proj['readiness'] = 'NOT STARTED'
                            
                    # C. Review Stage (Review Site Survey or general review)
                    elif 'Review Site Survey Information' in subitem_status or 'Review information provided' in subitem_status:
                        survey_status = subitem_status.get('Review Site Survey Information', 'Not Started')
                        info_status = subitem_status.get('Review information provided', 'Not Started')
                        if survey_status == 'Done' or info_status == 'Done':
                            proj['readiness'] = 'READY'
                        elif survey_status in ['Working on it', 'Stuck', 'Waiting on Info'] or info_status in ['Working on it', 'Stuck', 'Waiting on Info']:
                            proj['readiness'] = 'MISSING INFO'
                        else:
                            proj['readiness'] = 'NOT STARTED'
                            
                    # D. Pre-Design Stage
                    elif 'Send for Pre-Design' in subitem_status:
                        send_status = subitem_status.get('Send for Pre-Design', 'Not Started')
                        if send_status == 'Done':
                            proj['readiness'] = 'READY'
                        elif send_status in ['Working on it', 'Stuck', 'Waiting on Info']:
                            proj['readiness'] = 'MISSING INFO'
                        else:
                            proj['readiness'] = 'NOT STARTED'
                    else:
                        proj['readiness'] = 'NOT STARTED'
                            
                    group_data['projects'].append(proj)
                
                formatted_groups.append(group_data)
                
            return jsonify({'groups': formatted_groups})
        else:
            return jsonify({'error': f"Monday.com returned HTTP {resp.status_code}"}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# API Route: Fetch Outlook Calendar Events
@app.route('/api/outlook-calendar')
def get_outlook_calendar():
    token, err = get_outlook_token()
    if err:
        return jsonify({'error': err}), 400
        
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Query upcoming week
    now_str = requests.utils.quote(time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()))
    future_epoch = time.time() + (7 * 24 * 3600)  # 7 days future
    end_str = requests.utils.quote(time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(future_epoch)))
    
    url = f"https://graph.microsoft.com/v1.0/me/calendarView?startDateTime={now_str}&endDateTime={end_str}&$select=subject,start,end,location,organizer&$orderby=start/dateTime&$top=10"
    
    try:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            events = resp.json().get('value', [])
            formatted_events = []
            for ev in events:
                formatted_events.append({
                    'subject': ev.get('subject', 'No Title'),
                    'start': ev.get('start', {}).get('dateTime'),
                    'end': ev.get('end', {}).get('dateTime'),
                    'location': ev.get('location', {}).get('displayName', 'N/A'),
                    'organizer': ev.get('organizer', {}).get('emailAddress', {}).get('name', 'N/A')
                })
            return jsonify({'events': formatted_events})
        else:
            return jsonify({'error': f"Graph API returned HTTP {resp.status_code} - {resp.text}"}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# API Route: Fetch Outlook Unread Emails
@app.route('/api/outlook-unread-emails')
def get_outlook_unread_emails():
    token, err = get_outlook_token()
    if err:
        return jsonify({'error': err}), 400
        
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    # Query top 5 unread messages filtered by isRead eq false inside the Inbox folder for optimal speed and reliability
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=isRead eq false&$select=subject,from,receivedDateTime,bodyPreview&$orderby=receivedDateTime DESC&$top=5"
    
    try:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            messages = resp.json().get('value', [])
            formatted_emails = []
            for msg in messages:
                formatted_emails.append({
                    'subject': msg.get('subject', 'No Subject'),
                    'sender': msg.get('from', {}).get('emailAddress', {}).get('name', 'Unknown'),
                    'sender_email': msg.get('from', {}).get('emailAddress', {}).get('address', ''),
                    'received': msg.get('receivedDateTime'),
                    'preview': msg.get('bodyPreview', '')
                })
            return jsonify({'emails': formatted_emails})
        else:
            return jsonify({'error': f"Graph API returned HTTP {resp.status_code} - {resp.text}"}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Helper: Read reference standards files from local or sibling network_design_web directory
def get_reference_content(filename):
    current_dir = os.path.abspath(os.path.dirname(__file__))
    
    # 1. Try local references folder (copied for Render deployment)
    local_ref_path = os.path.join(current_dir, 'references', filename)
    if os.path.exists(local_ref_path):
        try:
            with open(local_ref_path, 'r') as f:
                return f.read()
        except Exception:
            pass
            
    # 2. Try finding in sibling network_design_web/references folder (for local dev)
    ref_path = os.path.abspath(os.path.join(current_dir, '..', 'network_design_web', 'references', filename))
    if os.path.exists(ref_path):
        try:
            with open(ref_path, 'r') as f:
                return f.read()
        except Exception:
            pass
            
    # 3. Direct sibling folder search fallback
    fallback_path = find_workspace_file(filename, os.path.join('network_design_web', 'references'))
    if os.path.exists(fallback_path):
        try:
            with open(fallback_path, 'r') as f:
                return f.read()
        except Exception:
            pass
            
    return ""

# API Route: Trigger Wi-Fi Design Engine using Pure Python Google GenAI SDK
@app.route('/api/generate-design/<property_code>')
def generate_design(property_code):
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        # Sibling folder load check fallback
        design_env = find_workspace_file('.env', 'network_design_web')
        if os.path.exists(design_env):
            try:
                with open(design_env, 'r') as f:
                    for line in f:
                        if 'GEMINI_API_KEY=' in line:
                            api_key = line.split('GEMINI_API_KEY=', 1)[1].strip().strip('"\'')
                            break
            except Exception:
                pass
                
    if not api_key:
        return jsonify({'error': 'GEMINI_API_KEY is not configured on the server. Please check your credentials.'}), 500

    brand_param = request.args.get('brand', '')
    sp_link = request.args.get('sp_link', '')
    brand = brand_param.upper().strip() if brand_param and brand_param.strip() else 'IHG'
    if brand == 'N/A' or brand == 'NONE':
        brand = 'IHG'
    
    # Map brand name to standard reference standards
    brand_file = 'ihg-meraki.md' # Default to IHG Meraki for unmapped flags
    if 'IHG' in brand:
        brand_file = 'ihg-meraki.md'
    elif 'WESTERN' in brand or 'BW' in brand:
        brand_file = 'best-western-aruba.md'
    elif 'CHOICE' in brand:
        brand_file = 'choice-aruba.md'
    elif 'WYNDHAM' in brand:
        brand_file = 'wyndham-aruba.md'
    elif 'OMADA' in brand:
        brand_file = 'independent-omada.md'
        
    # Build the final prompt for Gemini CLI
    prompt = (
        f"Activate the 'network-designer' skill. Create a complete, detailed Wi-Fi network design, "
        f"Bill of Materials (BOM) in Markdown table format, a Statement of Work (SOW), and a Logical Network Diagram "
        f"using Mermaid.js syntax for the property code '{property_code}'. Search SharePoint for files matching '{property_code}' "
    )
    if sp_link:
        prompt += f"(specifically at this folder location: '{sp_link}') "
        
    prompt += (
        f"to extract the true room counts, wall materials, and floorplans for the design. "
        f"Apply standard brand design parameters based on the reference file '{brand_file}'."
    )
    
    # Path for storing results
    output_dir = os.path.join(os.path.dirname(__file__), 'static', 'designs')
    os.makedirs(output_dir, exist_ok=True)
    
    md_file_path = os.path.join(output_dir, f"{property_code}_design_report.md")
    excel_file_path = os.path.join(output_dir, f"{property_code}_BOM.xlsx")
    
    try:
        # Run Gemini CLI in correct path environment using absolute path
        env = os.environ.copy()
        env['PATH'] = f"/opt/homebrew/bin:/usr/local/bin:{env.get('PATH', '')}"
        
        cmd = ["/opt/homebrew/bin/gemini", "-p", prompt, "--approval-mode", "yolo"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, env=env)
        
        raw_markdown = result.stdout
        if not raw_markdown:
            return jsonify({'error': "Gemini CLI returned an empty response"}), 500
            
        # Save raw report
        with open(md_file_path, "w") as f:
            f.write(raw_markdown)
            
        # Parse BOM table from markdown to generate Excel spreadsheet
        bom_rows = []
        is_table = False
        headers = []
        
        for line in raw_markdown.split('\n'):
            if '|' in line:
                parts = [p.strip() for p in line.split('|')[1:-1]]
                if not parts:
                    continue
                # Skip divider lines like |---|---|
                if all(all(c == '-' or c == ' ' for c in p) for p in parts):
                    continue
                if not is_table:
                    headers = parts
                    is_table = True
                else:
                    # Pad row if columns don't match headers
                    while len(parts) < len(headers):
                        parts.append('')
                    bom_rows.append(parts[:len(headers)])
            else:
                is_table = False
                
        # Generate Excel BOM if table was parsed
        if bom_rows and headers:
            try:
                import pandas as pd
                df = pd.DataFrame(bom_rows, columns=headers)
                df.to_excel(excel_file_path, index=False)
            except Exception as e:
                print(f"Error generating Excel for {property_code}: {e}")
                
        # Extract Mermaid diagram block
        mermaid_code = ""
        if "```mermaid" in raw_markdown:
            parts = raw_markdown.split("```mermaid", 1)
            if len(parts) == 2:
                mermaid_code = parts[1].split("```", 1)[0].strip()
                
        return jsonify({
            'success': True,
            'property_code': property_code,
            'report_md': raw_markdown,
            'mermaid_diagram': mermaid_code,
            'md_download_url': f"/static/designs/{property_code}_design_report.md",
            'excel_download_url': f"/static/designs/{property_code}_BOM.xlsx" if os.path.exists(excel_file_path) else None
        })
        
    except subprocess.CalledProcessError as e:
        err_msg = e.stderr or e.stdout or str(e)
        return jsonify({'error': f"Gemini CLI execution failed: {err_msg}"}), 500
    except Exception as e:
        return jsonify({'error': f"Gemini API execution failed: {str(e)}"}), 500

# API Route: Check SharePoint files for Design Readiness using Gemini CLI
@app.route('/api/check-readiness/<property_code>')
def check_readiness(property_code):
    sp_link = request.args.get('sp_link', '')
    
    prompt = (
        f"Activate the 'network-design-readiness' skill. Search SharePoint for files matching the property code '{property_code}'. "
    )
    if sp_link:
        prompt += f"(specifically at this folder location: '{sp_link}') "
        
    prompt += (
        f"Evaluate the discovered files for design readiness. List the files found and explain if they contain the floorplans, "
        f"wall materials, scales, and room counts needed to generate a professional Wi-Fi network design. "
        f"Conclude the report with a clear line: 'Overall Status: READY' if ready, or 'Overall Status: INCOMPLETE' if missing info."
    )
    
    try:
        # Run Gemini CLI using absolute path
        env = os.environ.copy()
        env['PATH'] = f"/opt/homebrew/bin:/usr/local/bin:{env.get('PATH', '')}"
        
        cmd = ["/opt/homebrew/bin/gemini", "-p", prompt, "--approval-mode", "yolo"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, env=env)
        
        raw_markdown = result.stdout
        if not raw_markdown:
            return jsonify({'error': "Gemini CLI returned an empty response"}), 500
            
        # Check if the output suggests it is ready
        is_ready = "Overall Status: READY" in raw_markdown or "Overall Status: Ready" in raw_markdown or "Status: READY" in raw_markdown
        
        return jsonify({
            'success': True,
            'property_code': property_code,
            'report_md': raw_markdown,
            'is_ready': is_ready
        })
        
    except subprocess.CalledProcessError as e:
        err_msg = e.stderr or e.stdout or str(e)
        return jsonify({'error': f"Gemini CLI execution failed: {err_msg}"}), 500
    except Exception as e:
        return jsonify({'error': f"Gemini Execution failed: {str(e)}"}), 500

# API Route: Mark Design Readiness subitem as Done on Monday.com
@app.route('/api/mark-ready/<property_code>', methods=['POST'])
def mark_ready(property_code):
    token = creds.get('MONDAY_API_TOKEN')
    board_id = creds.get('BOARD_ID')
    
    if not token:
        return jsonify({'error': 'Monday.com API Token is not configured.'}), 400
        
    url = "https://api.monday.com/v2"
    headers = {
        "Authorization": token,
        "Content-Type": "application/json"
    }
    
    # 1. Query subitems on board to locate 'Review information provided' or 'Review Site Survey Information' for the property
    query_subitems = """
    query {
      boards(ids: [%s]) {
        items_page(limit: 50) {
          items {
            id
            name
            subitems {
              id
              name
              board { id }
            }
          }
        }
      }
    }
    """ % board_id
    
    try:
        resp = requests.post(url, json={"query": query_subitems}, headers=headers)
        if resp.status_code == 200:
            data = resp.json()
            items = data.get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
            
            subitem_id = None
            subitem_board_id = None
            
            # Find matching item by property code prefix/substring
            for item in items:
                if property_code in item.get('name', ''):
                    subitems = item.get('subitems', []) or []
                    for s in subitems:
                        s_name = s.get('name', '')
                        if 'Review Site Survey Information' in s_name or 'Review information provided' in s_name:
                            subitem_id = s.get('id')
                            subitem_board_id = s.get('board', {}).get('id')
                            break
                    break
            
            if subitem_id and subitem_board_id:
                # Update status of subitem to Done
                mutation = """
                mutation {
                  change_simple_column_value (board_id: "%s", item_id: "%s", column_id: "status", value: "Done") {
                    id
                  }
                }
                """ % (subitem_board_id, subitem_id)
                
                m_resp = requests.post(url, json={"query": mutation}, headers=headers)
                if m_resp.status_code == 200:
                    return jsonify({'success': True, 'message': f'Successfully updated Monday.com subitem to Done!'})
                    
            return jsonify({'error': f"Could not find a matching survey/review subitem for {property_code} on Monday.com."}), 404
        else:
            return jsonify({'error': f"Monday.com returned HTTP {resp.status_code}: {resp.text}"}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# API Route: Fetch Weather via zero-config wttr.in JSON
@app.route('/api/weather')
def get_weather():
    city = request.args.get('city', 'Atlanta')
    url = f"https://wttr.in/{city}?format=j1"
    
    try:
        resp = requests.get(url)
        if resp.status_code == 200:
            data = resp.json()
            current_condition = data.get('current_condition', [{}])[0]
            nearest_area = data.get('nearest_area', [{}])[0]
            
            weather_data = {
                'temp_C': current_condition.get('temp_C', 'N/A'),
                'temp_F': current_condition.get('temp_F', 'N/A'),
                'desc': current_condition.get('weatherDesc', [{}])[0].get('value', 'N/A'),
                'humidity': current_condition.get('humidity', 'N/A'),
                'windspeed': current_condition.get('windspeedKmph', 'N/A'),
                'city': nearest_area.get('areaName', [{}])[0].get('value', city)
            }
            return jsonify(weather_data)
        else:
            return jsonify({'error': f"wttr.in returned HTTP {resp.status_code}"}), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    # We use 5001 to avoid default port conflicts
    app.run(host='0.0.0.0', port=port, debug=True)
