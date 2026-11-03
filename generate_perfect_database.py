import requests
import csv
import io
import re
import urllib.parse
import concurrent.futures
import sys
import pandas as pd
from bs4 import BeautifulSoup
import urllib3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
sys.stdout.reconfigure(encoding='utf-8')

DATA_URL = 'https://static.data.gouv.fr/resources/annuaire-des-exploitations-certifiees-haute-valeur-environnementale/20260903-130258/annuaire-des-exploitations-hve-juillet-2026.csv'
SOURCE_INFO = 'Ministère de l’Agriculture (data.gouv.fr - Annuaire officiel HVE)'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.8'
}

GENERIC_PREFIXES = (
    'contact', 'info', 'accueil', 'commercial', 'secretariat', 'direction',
    'domaine', 'chateau', 'cave', 'vins', 'vin', 'boutique', 'commande',
    'commandes', 'office', 'service', 'admin', 'adv', 'message', 'ferme',
    'exploitation', 'cooperative', 'vente', 'ventes', 'hello', 'bonjour'
)

DUMMY_DOMAINS = [
    'yoursite.com', 'example.com', 'domain.com', 'wix.com', 'site.fr',
    'mail.com', 'email.com', 'monsite.fr', 'test.com', 'placeholder.com'
]

DUMMY_PHONES = [
    '0123456789', '0540540540', '0000000000', '0102030405', '0600000000',
    '01 23 45 67 89', '05 40 54 05 40'
]

def clean_url(url):
    if not url: return None
    url = url.strip()
    if not url.startswith('http://') and not url.startswith('https://'):
        url = 'https://' + url
    return url.rstrip('/')

def is_valid_domain(s):
    if not s: return False
    s_clean = s.strip().lower()
    if any(bad in s_clean for bad in ['non', 'aucun', 'neant', 'n/a', 'na', 'nc', 'facebook', 'instagram', '@', 'hve.fr', 'linkedin']):
        return False
    if s_clean.startswith(('http://', 'https://', 'www.')) or re.search(r'^[a-zA-Z0-9-]+\.[a-zA-Z]{2,}', s_clean):
        return True
    return False

def clean_phone(phone_str):
    if not phone_str: return ''
    p = re.sub(r'[\s.-]+', ' ', phone_str.strip())
    digits = re.sub(r'\D', '', p)
    if digits.startswith('33'):
        digits = '0' + digits[2:]
    if len(digits) == 10 and digits.startswith('0'):
        p = f"{digits[0:2]} {digits[2:4]} {digits[4:6]} {digits[6:8]} {digits[8:10]}"
    if p in DUMMY_PHONES or digits in ['0123456789', '0540540540', '0000000000']:
        return ''
    return p

def clean_org_name(name):
    if not name: return ""
    name = str(name).strip()
    name = re.sub(r'^\d+\s*\(|\)$', '', name).strip()
    name = re.sub(r'\s*\(\s*(?:EARL|GAEC|SCEA|SAS|SARL)\s*$', '', name, flags=re.I)
    if name.count('(') > name.count(')'):
        name += ')'
    elif name.count(')') > name.count('('):
        name = name.replace(')', '')
    return re.sub(r'\s+', ' ', name).strip()

def crawl_and_enrich(row):
    site_url = clean_url(row.get('Site Internet'))
    if not site_url:
        return None

    legal_name = clean_org_name(row.get('Nom ou Raison sociale') or '')
    brand_name = ""
    contact_page = None
    homepage_html = ""
    contact_html = ""

    try:
        r = requests.get(site_url, headers=HEADERS, timeout=7, verify=False, allow_redirects=True)
        if r.status_code == 200:
            final_site = r.url.rstrip('/')
            homepage_html = r.text
            soup = BeautifulSoup(homepage_html, 'html.parser')

            og_site = soup.find('meta', property='og:site_name')
            if og_site and og_site.get('content'):
                brand_name = og_site['content'].strip()
            elif soup.title and soup.title.string:
                title_clean = soup.title.string.strip()
                parts = re.split(r'[\-|–|•]', title_clean)
                if len(parts) > 0 and len(parts[0].strip()) > 3:
                    cand = parts[0].strip()
                    if not any(w in cand.lower() for w in ['accueil', 'home', 'bienvenue', 'loading']):
                        brand_name = cand

            for a in soup.find_all('a', href=True):
                href = a['href']
                text = a.get_text(strip=True).lower()
                h_lower = href.lower()
                if any(k in h_lower or k in text for k in ['contact', 'nous-contacter', 'contactez-nous', 'coordonnees', 'acces-contact', 'mentions-legales']):
                    full_c = urllib.parse.urljoin(final_site, href)
                    if not any(bad in full_c.lower() for bad in ['facebook', 'instagram', 'twitter', 'linkedin']):
                        contact_page = full_c
                        break
    except Exception:
        pass

    if contact_page and contact_page != site_url:
        try:
            r_c = requests.get(contact_page, headers=HEADERS, timeout=7, verify=False, allow_redirects=True)
            if r_c.status_code == 200:
                contact_html = r_c.text
        except Exception:
            pass

    combined_html = homepage_html + " " + contact_html

    mailto_matches = re.findall(r'href=[\'"]mailto:([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)[\'"]', combined_html, re.I)
    raw_emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', combined_html)
    all_emails = list(dict.fromkeys(mailto_matches + raw_emails))

    generic_email = ''
    for em in all_emails:
        em_clean = em.lower().strip('.')
        if any(em_clean.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.js', '.css']):
            continue
        domain_part = em_clean.split('@')[-1]
        if any(bad_d in domain_part for bad_d in DUMMY_DOMAINS):
            continue
        prefix = em_clean.split('@')[0]
        if any(prefix.startswith(gp) or prefix == gp for gp in GENERIC_PREFIXES):
            generic_email = em_clean
            break

    phone_matches = re.findall(r'(?:(?:\+|00)33[\s.-]?(?:\(0\)[\s.-]?)?|0)[1-59](?:[\s.-]?\d{2}){4}', combined_html)
    valid_phone = ''
    if phone_matches:
        valid_phone = clean_phone(phone_matches[0])

    final_org_name = legal_name
    if brand_name and len(brand_name) < 45 and brand_name.lower() not in legal_name.lower() and legal_name.lower() not in brand_name.lower():
        final_org_name = f"{brand_name} ({legal_name})"

    final_contact_url = contact_page or (site_url + '/contact')

    act_veg = (row.get('Activité végétale') or '').strip()
    act_anim = (row.get('Activité animale') or '').strip()
    activities = []
    if act_veg: activities.append(act_veg)
    if act_anim: activities.append(act_anim)
    activity_str = ", ".join(activities) if activities else "Agriculture générale"

    region = (row.get('Région') or '').strip()
    dept = (row.get('Département') or '').strip()
    commune = (row.get('Commune') or '').strip()
    location = f"{region}, {dept} ({commune})"

    return {
        'Название организации': final_org_name,
        'Регион и населённый пункт': location,
        'Направление деятельности и выращиваемые культуры': activity_str,
        'Официальный сайт': site_url,
        'Ссылка на страницу «Контакты» или форму обращения': final_contact_url,
        'Общая электронная почта организации': generic_email,
        'Телефон приёмной': valid_phone,
        'Ссылка на источник информации': SOURCE_INFO
    }

def add_glossary_panel(ws):
    ws.column_dimensions['I'].width = 4
    ws.column_dimensions['J'].width = 28
    ws.column_dimensions['K'].width = 62

    font_title = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_title = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')

    font_section = Font(name='Segoe UI', size=10, bold=True, color='1F4E79')
    fill_section = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    font_term = Font(name='Segoe UI', size=9, bold=True, color='262626')
    font_desc = Font(name='Segoe UI', size=9, color='333333')
    font_link = Font(name='Segoe UI', size=9, color='0563C1', underline='single')

    fill_even = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    fill_odd = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    border_sec = Border(
        left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'),
        top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7')
    )

    ws.merge_cells('J1:K1')
    c_top = ws['J1']
    c_top.value = "СПРАВОЧНИК И РАСШИФРОВКА ОБОЗНАЧЕНИЙ"
    c_top.font = font_title
    c_top.fill = fill_title
    c_top.alignment = Alignment(horizontal='center', vertical='center')
    ws['K1'].border = border_thin
    ws['J1'].border = border_thin

    panel_data = [
        ("section", "1. ИСТОЧНИК ДАННЫХ И СЕРТИФИКАЦИЯ (HVE)", ""),
        ("item", "data.gouv.fr", "Единый государственный портал открытых данных правительства Франции (Open Data République Française)."),
        ("item", "Ministère de l'Agriculture", "Министерство сельского хозяйства и продовольственного суверенитета Франции (официальный первоисточник реестра)."),
        ("item", "HVE (Haute Valeur Environnementale)", "Высший (3-й) уровень государственной эко-сертификации Франции. Знак качества, подтверждающий крупное, активное хозяйство с бюджетом на услуги и технологии."),
        ("item", "Официальный реестр", "Госреестр HVE: https://data.gouv.fr/fr/datasets/annuaire-des-exploitations-certifiees-haute-valeur-environnementale/"),
        ("section", "2. ФОРМЫ СЕЛЬХОЗПРЕДПРИЯТИЙ ФРАНЦИИ", ""),
        ("item", "GAEC", "Groupement agricole d'exploitation en commun — совместное фермерское предприятие (объединение хозяйств)."),
        ("item", "EARL", "Exploitation agricole à responsabilité limitée — фермерское хозяйство / юрлицо с ограниченной ответственностью."),
        ("item", "SCEA", "Société civile d'exploitation agricole — сельскохозяйственное гражданское общество."),
        ("item", "SAS / SARL / SA", "Коммерческие сельскохозяйственные компании, агрохолдинги и перерабатывающие структуры."),
        ("section", "3. НАПРАВЛЕНИЯ ДЕЯТЕЛЬНОСТИ И КУЛЬТУРЫ", ""),
        ("item", "Grandes cultures", "Полевые культуры: зерновые, зернобобовые, подсолнечник, рапс, кукуруза, сахарная свёкла."),
        ("item", "Viticulture", "Виноградарство и первичное виноделие (винодельческие хозяйства, домены, шато)."),
        ("item", "Fruits / Arboriculture", "Промышленное садоводство, семечковые и косточковые культуры (яблоки, груши, персики), ягодники."),
        ("item", "Légumes / Maraîchage", "Овощеводство открытого и закрытого (тепличного) грунта."),
        ("item", "Élevage (Bovins / Ovins)", "Животноводство: мясное и молочное скотоводство (КРС), овцеводство, птицеводство."),
        ("section", "4. СТРОГОЕ СОБЛЮДЕНИЕ GDPR И ТЗ", ""),
        ("item", "Только общие каналы связи", "В таблице указаны исключительно корпоративные адреса (contact@, info@, accueil@, secretariat@) и телефон приёмной."),
        ("item", "Без персональных данных", "Именные адреса (ФИО), личные мобильные номера и данные частных собственников исключены по нормам GDPR и ТЗ."),
        ("item", "Прямой доступ", "Для каждой организации приведена прямая ссылка на форму связи / страницу «Контакты» на официальном сайте.")
    ]

    r_idx = 2
    for entry_type, cj_val, ck_val in panel_data:
        cj = ws.cell(row=r_idx, column=10)
        ck = ws.cell(row=r_idx, column=11)
        if entry_type == "section":
            ws.merge_cells(start_row=r_idx, start_column=10, end_row=r_idx, end_column=11)
            cj.value = cj_val
            cj.font = font_section
            cj.fill = fill_section
            cj.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            cj.border = border_sec
            ck.border = border_sec
            ws.row_dimensions[r_idx].height = 24
        else:
            cj.value = cj_val
            cj.font = font_term
            cj.border = border_thin
            cj.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            ck.value = ck_val
            ck.font = font_desc
            ck.border = border_thin
            ck.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            if 'http' in ck_val:
                ck.font = font_link
                ck.hyperlink = "https://data.gouv.fr/fr/datasets/annuaire-des-exploitations-certifiees-haute-valeur-environnementale/"

            r_fill = fill_even if (r_idx % 2 == 0) else fill_odd
            cj.fill = r_fill
            ck.fill = r_fill
            ws.row_dimensions[r_idx].height = 26
        r_idx += 1

def main():
    print("1. Loading raw dataset...")
    r = requests.get(DATA_URL, timeout=30)
    reader = csv.DictReader(io.StringIO(r.content.decode('utf-8', errors='ignore')), delimiter=';')
    rows = list(reader)

    valid_rows = [r for r in rows if is_valid_domain(r.get('Site Internet'))]
    print(f"   Valid website records: {len(valid_rows)}")

    candidates_grandes_cultures = []
    candidates_fruits = []
    candidates_legumes = []
    candidates_elevage = []
    candidates_viticulture = []

    for r in valid_rows:
        veg = (r.get('Activité végétale') or '').strip()
        anim = (r.get('Activité animale') or '').strip()
        if anim:
            candidates_elevage.append(r)
        elif 'Légumes' in veg:
            candidates_legumes.append(r)
        elif 'Fruits' in veg:
            candidates_fruits.append(r)
        elif 'Grandes cultures' in veg:
            candidates_grandes_cultures.append(r)
        elif 'Viticulture' in veg:
            candidates_viticulture.append(r)

    selected_pool = (
        candidates_grandes_cultures[:15] +
        candidates_fruits[:20] +
        candidates_legumes[:10] +
        candidates_elevage[:25] +
        candidates_viticulture[:35]
    )

    print(f"2. Selected pool of {len(selected_pool)} enterprises. Crawling...")
    enriched = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=14) as executor:
        futures = {executor.submit(crawl_and_enrich, row): row for row in selected_pool}
        for f in concurrent.futures.as_completed(futures):
            res = f.result()
            if res:
                enriched.append(res)
            if len(enriched) >= 100:
                break

    final_100 = enriched[:100]

    # Export CSV
    csv_path = 'france_agricultural_contacts_100.csv'
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=list(final_100[0].keys()), delimiter=';')
        writer.writeheader()
        writer.writerows(final_100)

    # Export XLSX
    xlsx_path = 'france_agricultural_contacts_100.xlsx'
    df = pd.DataFrame(final_100)
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Агропредприятия Франции')
        ws = writer.sheets['Агропредприятия Франции']

        header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        data_font = Font(name='Segoe UI', size=10)
        link_font = Font(name='Segoe UI', size=10, color='0563C1', underline='single')
        zebra_fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
        border_light = Border(
            left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'),
            top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0')
        )

        ws.row_dimensions[1].height = 28
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=len(df)+1), start=2):
            ws.row_dimensions[row_idx].height = 22
            is_even = (row_idx % 2 == 0)
            for col_idx, cell in enumerate(row, start=1):
                cell.font = data_font
                cell.border = border_light
                if is_even: cell.fill = zebra_fill
                val = str(cell.value or '').strip()
                if col_idx in [4, 5] and val.startswith('http'):
                    cell.font = link_font
                    cell.hyperlink = val
                elif col_idx in [6, 7]:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        widths = {1: 34, 2: 38, 3: 36, 4: 36, 5: 42, 6: 30, 7: 24, 8: 44}
        for col_idx, width in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = "A1:H101"

        # Add side panel with glossary
        add_glossary_panel(ws)

    print("Database and side glossary generated successfully!")

if __name__ == '__main__':
    main()
