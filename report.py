import requests
import base64
import re
import os
import json
from datetime import datetime, timedelta, timezone
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sendgrid
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64 as b64

# ── Credentials from GitHub Secrets ───────────────────────
USERNAME = os.environ['CRM_USERNAME']
PASSWORD = os.environ['CRM_PASSWORD']
SENDGRID_API_KEY = os.environ['SENDGRID_API_KEY']
TO_EMAIL = os.environ.get('TO_EMAIL', 'hs@jfrecycle.com')
FROM_EMAIL = os.environ.get('FROM_EMAIL', 'hs@jfrecycle.com')

BASE_URL = 'https://prophetOnDemand.com/prophet/prophetwebservices/AvtProphetApi/odata'
EXCEL_MAX = 32000

token = base64.b64encode(f'{USERNAME}:{PASSWORD}'.encode()).decode()
HEADERS = {'Authorization': f'Basic {token}', 'Accept': 'application/json;odata=verbose'}

cutoff_dt = datetime.now(timezone.utc) - timedelta(days=7)
cutoff    = cutoff_dt.strftime('%Y-%m-%dT00:00:00')
print(f'Pulling notes from last 7 days (since {cutoff[:10]})')

def safe_str(val):
    if val is None: return ''
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', str(val))
    return (s[:EXCEL_MAX] + '... [truncated]') if len(s) > EXCEL_MAX else s

def parse_note(raw):
    if not raw: return '', '', '', ''
    text   = raw.replace('\r\n', '\n').replace('\r', '\n')
    lines  = text.split('\n')
    header = lines[0]
    body   = '\n'.join(lines[1:]).strip()
    chunks = [c.strip() for c in re.split(r'-{3,}', header) if c.strip()]
    date_author   = chunks[0] if chunks else ''
    activity_type = chunks[1] if len(chunks) > 1 else ''
    m_author = re.search(r'Modified by:\s*(.+)$', date_author, re.IGNORECASE)
    author   = m_author.group(1).strip() if m_author else ''
    m_date   = re.match(r'^(.+?)\s*-\s*Modified by:', date_author, re.IGNORECASE)
    date_str = m_date.group(1).strip() if m_date else ''
    return author, activity_type, date_str, body

def week_start(dt):
    return dt - timedelta(days=dt.weekday())

# ── Fetch notes ────────────────────────────────────────────
note_filter = f"CreatedDate ge datetime'{cutoff}'"
all_notes, skip = [], 0
while True:
    url = (f'{BASE_URL}/Notes?$top=100&$skip={skip}'
           f'&$filter={requests.utils.quote(note_filter)}&$orderby=CreatedDate desc')
    resp = requests.get(url, headers=HEADERS)
    if resp.status_code == 401: raise Exception('Login failed.')
    if resp.status_code != 200: raise Exception(f'API error {resp.status_code}: {resp.text[:300]}')
    page = resp.json().get('value') or resp.json().get('d', {}).get('results', [])
    if not page: break
    all_notes.extend(page)
    print(f'  ...{len(all_notes)} notes fetched')
    if len(page) < 100: break
    skip += 100

print(f'Total notes: {len(all_notes)}')

# ── Process notes ──────────────────────────────────────────
weekly_activity = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
opp_cache = {}

def get_opp(entity_id):
    if entity_id in opp_cache: return opp_cache[entity_id]
    try:
        r = requests.get(f"{BASE_URL}/OpportunityViews?$filter=Id eq guid'{entity_id}'", headers=HEADERS)
        if r.status_code == 200:
            items = r.json().get('value') or r.json().get('d', {}).get('results', [])
            if items:
                opp_cache[entity_id] = items[0]
                return items[0]
    except: pass
    opp_cache[entity_id] = None
    return None

results = []
for i, n in enumerate(all_notes):
    entity_id = n.get('EntityId', '')
    raw       = n.get('Text') or ''
    author, activity_type, date_str, body = parse_note(raw)
    if not author and not body: continue
    if author and author.lower() == 'system': continue
    opp      = get_opp(entity_id)
    company  = safe_str(opp.get('CompanyName', '')) if opp else ''
    opp_name = safe_str(opp.get('RecordDescription', '')) if opp else ''
    stage    = safe_str(opp.get('UserDefined02', '')) if opp else ''
    city     = safe_str(opp.get('CompanyBusinessCity', '')) if opp else ''
    state    = safe_str(opp.get('CompanyBusinessState', '')) if opp else ''
    compete  = safe_str(opp.get('UserDefined34', '')) if opp else ''
    print(f'  [{i+1}/{len(all_notes)}] {author} — {activity_type} — {company}')
    api_date  = (n.get('CreatedDate') or '')[:10]
    backdated = False
    if date_str and api_date:
        try:
            header_date = datetime.strptime(date_str.strip().split()[0], '%m/%d/%Y').strftime('%Y-%m-%d')
            backdated   = header_date != api_date
        except: pass
    if author and activity_type and api_date:
        try:
            note_dt = datetime.fromisoformat(api_date)
            ws      = week_start(note_dt)
            weekly_activity[ws][author][activity_type] += 1
        except: pass
    results.append({
        'Company Name':     company,
        'Opportunity Name': opp_name,
        'Stage':            stage,
        'City':             city,
        'State':            state,
        'Competition':      compete,
        'User':             safe_str(author),
        'Date':             safe_str(date_str),
        'Activity Type':    safe_str(activity_type),
        'Note':             safe_str(body),
        '_backdated':       backdated,
        '_api_date':        api_date,
    })

print(f'Processed {len(results)} note rows')

# ── Build Excel ────────────────────────────────────────────
wb = Workbook()
hdr_fill      = PatternFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
hdr_font      = Font(bold=True, color='FFFFFF', size=11)
total_fill    = PatternFill(start_color='DBEAFE', end_color='DBEAFE', fill_type='solid')
total_font    = Font(bold=True, size=11)
flag_fill     = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
flag_hdr_fill = PatternFill(start_color='FFC107', end_color='FFC107', fill_type='solid')
wrap          = Alignment(vertical='top', wrap_text=True)
mid           = Alignment(horizontal='center', vertical='center')
left_mid      = Alignment(horizontal='left', vertical='center')
thin          = Side(style='thin', color='D1D5DB')
border        = Border(left=thin, right=thin, top=thin, bottom=thin)
alt_fill      = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')

# Sheet 1 - Notes by Activity
ws1 = wb.active
ws1.title = 'Notes by Activity'
headers         = ['Company Name','Opportunity Name','Stage','City','State','Competition','User','Date','Activity Type','Note']
display_headers = headers + ['Backdated?']
for ci, h in enumerate(display_headers, 1):
    c = ws1.cell(row=1, column=ci, value=h)
    c.font=hdr_font; c.alignment=mid; c.border=border
    c.fill = flag_hdr_fill if h == 'Backdated?' else hdr_fill
for ri, rec in enumerate(results, 2):
    is_backdated = rec.get('_backdated', False)
    row_fill     = flag_fill if is_backdated else (alt_fill if ri % 2 == 0 else None)
    for ci, key in enumerate(headers, 1):
        c = ws1.cell(row=ri, column=ci, value=rec[key])
        c.border=border; c.alignment=wrap
        if row_fill: c.fill=row_fill
    c = ws1.cell(row=ri, column=len(display_headers),
                 value=f'YES — entered {rec["_api_date"]}' if is_backdated else '')
    c.border=border; c.alignment=mid
    if is_backdated:
        c.fill=flag_fill; c.font=Font(bold=True, color='856404', size=11)
    elif row_fill: c.fill=row_fill
cw = {'Company Name':28,'Opportunity Name':32,'Stage':26,'City':18,'State':10,'Competition':22,
      'User':22,'Date':22,'Activity Type':28,'Note':70,'Backdated?':28}
for ci, h in enumerate(display_headers, 1):
    ws1.column_dimensions[get_column_letter(ci)].width = cw.get(h, 20)
ws1.row_dimensions[1].height = 22
for ri in range(2, len(results)+2): ws1.row_dimensions[ri].height = 80
ws1.freeze_panes='A2'; ws1.auto_filter.ref=ws1.dimensions

# Sheet 2 - Weekly Activity Summary
ws2 = wb.create_sheet(title='Activity Summary (Weekly)')
all_types = sorted(set(at for wd in weekly_activity.values() for ud in wd.values() for at in ud))
all_weeks = sorted(weekly_activity.keys())
sh = ['Week', 'User'] + all_types + ['TOTAL']
for ci, h in enumerate(sh, 1):
    c = ws2.cell(row=1, column=ci, value=h)
    c.font=hdr_font; c.fill=hdr_fill; c.alignment=mid; c.border=border
ri = 2
for ws_dt in all_weeks:
    week_end   = ws_dt + timedelta(days=6)
    week_label = f"{ws_dt.strftime('%m/%d/%Y')} - {week_end.strftime('%m/%d/%Y')}"
    week_users = sorted(weekly_activity[ws_dt].keys())
    for user in week_users:
        is_alt    = ri % 2 == 0
        row_total = 0
        c = ws2.cell(row=ri, column=1, value=week_label)
        c.font=Font(size=11); c.alignment=left_mid; c.border=border
        if is_alt: c.fill=alt_fill
        c = ws2.cell(row=ri, column=2, value=user)
        c.font=Font(bold=True, size=11); c.alignment=left_mid; c.border=border
        if is_alt: c.fill=alt_fill
        for ci, at in enumerate(all_types, 3):
            count = weekly_activity[ws_dt][user].get(at, 0)
            row_total += count
            c = ws2.cell(row=ri, column=ci, value=count if count > 0 else '')
            c.alignment=mid; c.border=border
            if is_alt: c.fill=alt_fill
        c = ws2.cell(row=ri, column=len(sh), value=row_total)
        c.font=total_font; c.alignment=mid; c.border=border; c.fill=total_fill
        ri += 1
    # Week subtotal
    c = ws2.cell(row=ri, column=1, value=f'WEEK TOTAL — {week_label}')
    c.font=total_font; c.fill=total_fill; c.alignment=left_mid; c.border=border
    ws2.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=2)
    week_grand = 0
    for ci, at in enumerate(all_types, 3):
        wt = sum(weekly_activity[ws_dt][u].get(at, 0) for u in week_users)
        week_grand += wt
        c = ws2.cell(row=ri, column=ci, value=wt if wt > 0 else '')
        c.font=total_font; c.fill=total_fill; c.alignment=mid; c.border=border
    c = ws2.cell(row=ri, column=len(sh), value=week_grand)
    c.font=total_font; c.fill=total_fill; c.alignment=mid; c.border=border
    ri += 1
ws2.column_dimensions['A'].width = 32
ws2.column_dimensions['B'].width = 26
for ci in range(3, len(sh)+1):
    ws2.column_dimensions[get_column_letter(ci)].width = 26
ws2.row_dimensions[1].height = 22
for r in range(2, ri): ws2.row_dimensions[r].height = 22
ws2.freeze_panes='C2'

filename = f'avidian_weekly_{datetime.now(timezone.utc).strftime("%Y-%m-%d")}.xlsx'
wb.save(filename)
print(f'Saved {filename}')

# ── Send email via SendGrid ────────────────────────────────
with open(filename, 'rb') as f:
    file_data = f.read()
encoded = b64.b64encode(file_data).decode()

today     = datetime.now(timezone.utc).strftime('%B %d, %Y')
week_ago  = (datetime.now(timezone.utc) - timedelta(days=7)).strftime('%B %d, %Y')

message = Mail(
    from_email=FROM_EMAIL,
    to_emails=TO_EMAIL,
    subject=f'Avidian CRM Weekly Report — {week_ago} to {today}',
    html_content=f'''
    <h2>Avidian CRM Weekly Report</h2>
    <p>Your automated weekly CRM report is attached.</p>
    <p><strong>Period:</strong> {week_ago} to {today}</p>
    <p><strong>Total activities logged:</strong> {len(results)}</p>
    <p style="color:#856404;"><strong>Note:</strong> Rows highlighted in yellow indicate backdated entries.</p>
    <br>
    <p style="color:#888; font-size:12px;">This report was generated automatically every Monday at 7:30am EST.</p>
    '''
)

attachment = Attachment(
    FileContent(encoded),
    FileName(filename),
    FileType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    Disposition('attachment')
)
message.attachment = attachment

sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
response = sg.send(message)
print(f'Email sent! Status: {response.status_code}')

# ── Duplicate Touch Detection ──────────────────────────────
from datetime import datetime as dt

touch_events = []
for rec in results:
    api_date = rec.get('_api_date', '')
    raw_date = rec.get('Date', '')
    user     = rec.get('User', '')
    company  = rec.get('Company Name', '')
    atype    = rec.get('Activity Type', '')
    if not user or not company or not api_date: continue
    parsed_dt = None
    for fmt in ['%m/%d/%Y %I:%M %p', '%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y']:
        try:
            parsed_dt = dt.strptime(raw_date.strip(), fmt)
            break
        except: pass
    if parsed_dt is None:
        try: parsed_dt = dt.fromisoformat(api_date)
        except: continue
    touch_events.append((parsed_dt, user, company, atype))

touch_events.sort(key=lambda x: (x[1], x[2], x[0]))

duplicates = []
i = 0
while i < len(touch_events):
    first_dt, user, company, atype = touch_events[i]
    group = [touch_events[i]]
    j = i + 1
    while j < len(touch_events):
        next_dt, next_user, next_company = touch_events[j][0], touch_events[j][1], touch_events[j][2]
        if next_user != user or next_company != company: break
        diff_mins = (next_dt - first_dt).total_seconds() / 60
        if diff_mins <= 30:
            group.append(touch_events[j])
            j += 1
        else: break
    if len(group) > 1:
        for dup in group[1:]:
            duplicates.append((first_dt, dup[0], user, company, dup[3]))
    i = j if j > i + 1 else i + 1

# Write duplicate section to sheet 2
gap = 3
dup_start_row = ri + gap

c = ws2.cell(row=dup_start_row, column=1, value='DUPLICATE TOUCHES (same user, same company, within 30 mins)')
c.font=Font(bold=True, color='FFFFFF', size=12)
c.fill=PatternFill(start_color='B45309', end_color='B45309', fill_type='solid')
c.alignment=left_mid; c.border=border
ws2.merge_cells(start_row=dup_start_row, start_column=1, end_row=dup_start_row, end_column=6)

dup_headers = ['User', 'Company', 'First Touch', 'Duplicate Touch', 'Time Apart (mins)', 'Activity Type']
for ci, h in enumerate(dup_headers, 1):
    c = ws2.cell(row=dup_start_row+1, column=ci, value=h)
    c.font=hdr_font; c.fill=hdr_fill; c.alignment=mid; c.border=border

dup_fill = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
dup_alt  = PatternFill(start_color='FDE68A', end_color='FDE68A', fill_type='solid')

if not duplicates:
    c = ws2.cell(row=dup_start_row+2, column=1, value='No duplicate touches found in this period.')
    c.font=Font(italic=True, size=11); c.alignment=left_mid
    ws2.merge_cells(start_row=dup_start_row+2, start_column=1, end_row=dup_start_row+2, end_column=6)
else:
    for di, (first_dt, dup_dt, user, company, atype) in enumerate(duplicates, dup_start_row+2):
        fill = dup_fill if di % 2 == 0 else dup_alt
        mins_apart = round((dup_dt - first_dt).total_seconds() / 60, 1)
        row_vals = [user, company,
                    first_dt.strftime('%m/%d/%Y %I:%M %p'),
                    dup_dt.strftime('%m/%d/%Y %I:%M %p'),
                    mins_apart, atype]
        for ci, val in enumerate(row_vals, 1):
            c = ws2.cell(row=di, column=ci, value=val)
            c.border=border; c.fill=fill
            c.alignment=mid if ci > 2 else left_mid
        ws2.row_dimensions[di].height = 20

ws2.column_dimensions['A'].width = max(32, ws2.column_dimensions['A'].width)
ws2.column_dimensions['B'].width = 30
ws2.column_dimensions['C'].width = 22
ws2.column_dimensions['D'].width = 22
ws2.column_dimensions['E'].width = 20
ws2.column_dimensions['F'].width = 28

print(f'Found {len(duplicates)} duplicate touches')

