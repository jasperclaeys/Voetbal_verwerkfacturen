import pdfplumber
import re
import os
import pandas as pd
import csv
from collections import defaultdict
import imaplib
import email
from email.header import decode_header
from dotenv import load_dotenv
from datetime import datetime
import traceback


#Gloabals
base_dir = os.getcwd()
load_dotenv(dotenv_path="password.env")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

def verwerkfactuur(pdf_file,pdf_naam):
    totaal_per_cat = defaultdict(float)
    bedrag_per_weborder=defaultdict(float)
    with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    lines = full_text.splitlines()
    factuurnummer=vind_factuurnummer(lines)
    vervaldag =vind_vervaldag(lines)
    Bedrag_BTW_EXCL=vind_bedrag_totaal_zonder_btw(lines)
    print("Bedrag BTW excl is "+ str(Bedrag_BTW_EXCL))
    Bedrag_BTW_incl=vind_bedrag_inclusief_btw(lines)
    bedrag_per_weborder=totaalbedrag_per_weborder(lines)
    totaal_per_cat =totaal_per_categorie(bedrag_per_weborder,base_dir+"/orders.csv")
    print(totaal_per_cat.values())
    weborder_b= check_gelijkheid("bedragen per weborder", Bedrag_BTW_EXCL, sum(bedrag_per_weborder.values()))
    categorie_b =check_gelijkheid("bedragen per categorie", Bedrag_BTW_EXCL, sum(totaal_per_cat.values()))
    if weborder_b and categorie_b:
        print("✅ Alles klopt!")
        schrijf_factuurregel(factuurnummer,vervaldag,Bedrag_BTW_EXCL,Bedrag_BTW_incl,totaal_per_cat,pdf_naam)    
        return (
        factuurnummer,
        vervaldag,
        Bedrag_BTW_EXCL,
        Bedrag_BTW_incl,
        totaal_per_cat.get("CAT1", 0.0),
        totaal_per_cat.get("CAT2", 0.0),
        totaal_per_cat.get("CAT3", 0.0),
        totaal_per_cat.get("CAT4", 0.0),
        sum(totaal_per_cat.values()))
    else:
        print("❌ Er is een probleem met de bedragen! Dit op factuurnummer:",factuurnummer,pdf_file)
        return None
def check_gelijkheid(label, bedrag1, bedrag2, marge=0.01):
    b1 = round(bedrag1, 2)
    b2 = round(bedrag2, 2)
    verschil = round(b2 - b1, 2)

    if abs(verschil) <= marge:
        print(f"✅ {label}: bedragen komen overeen (±€{marge:.2f}) → €{b1:.2f} vs €{b2:.2f}")
        return True
    else:
        print(f"❌ {label}: afwijking buiten marge (±€{marge:.2f})")
        print(f"   ⚠️ Verschil: €{verschil:+.2f} → €{b1:.2f} vs €{b2:.2f}")
        return False
def vind_vervaldag(lines):
    for line in lines:
        match = re.search(r'Vervaldag\s+(\d{2}-[A-Za-z]{3}-\d{4})', line)
        if match:
            try:
                vervaldatum = datetime.strptime(match.group(1), "%d-%b-%Y").date()
                return vervaldatum
            except ValueError:
                print(f"⚠️ Ongeldig datumformaat: {match.group(1)}")
    return None  # Geen vervaldag gevonden
def vind_factuurnummer(lines):
    for line in lines:
        match = re.search(r'Factuurnr\.\s*(\d{2}-\d{4}-\d{5})', line)
        if match:
            return match.group(1)
    return None  # Geen factuurnummer gevonden
def vind_bedrag_totaal_zonder_btw(lines):
    for line in lines:
        if "Betalingsinstructies: Totaal zonder BTW" in line:
            match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
            if match:
                print (line)
                return parse_bedrag_europees(match.group(1))
    return 0.0  # Geen match gevonden
def vind_bedrag_inclusief_btw(lines):
    for line in lines:
        if "Te betalen EUR" in line:
            match = re.search(r'(\d+,\d{2})', line)
            if match:
                return parse_bedrag_europees(match.group(1))
    return 0.0  # Geen match gevonden
def vind_weborders_met_posities(lines):
    resultaten = []
    for i, line in enumerate(lines):
        matches = re.findall(r'(?:Website Order)\s*(\d+)', line)
        for match in matches:
            resultaten.append((match, i))  # (weborder, positie)
    return resultaten
def totaalbedrag_per_weborder(lines):
    """
    Doorloopt een lijst van regels (lines) en berekent per weborder het totaalbedrag van geldige bestellingen.
    Een bestelling is geldig als:
    - De regel bevat een bedrag aan het einde (zoals '34,95')
    - Voldoet aan de controle: derde cijfer = eerste × tweede

    Returns:
        dict: weborder → totaalbedrag (float)
    """
    weborder_posities = vind_weborders_met_posities(lines)

    totaal_per_weborder = defaultdict(float)

    # Verwerk elk weborderblok afzonderlijk
    for idx in range(len(weborder_posities) - 1):
        weborder, start = weborder_posities[idx]
        _, end = weborder_posities[idx + 1]

        # Doorloop alle regels binnen dit weborderblok
        for line in lines[start:end]:
            stripped = line.strip()
            # Regel eindigt op drie getallen waarvan eerste × tweede ≈ derde
            match = re.search(r"(\d+)\s+([\d.,]+)\s+([\d.,]+)$", stripped)
            if match:
                qty = int(match.group(1))
                
                unit = parse_bedrag_europees(match.group(2))
                
                total = parse_bedrag_europees(match.group(3))
                
                if abs(qty * unit - total) < 0.01:
                    totaal_per_weborder[weborder] += total
                   
            

    print(totaal_per_weborder)
    return totaal_per_weborder
def totaal_per_categorie(orderbedragen, pad_csv):
    """
    Verwerkt een CSV-bestand met weborder → categorie mapping en telt per categorie
    het totaalbedrag op uit de meegegeven orderbedragen.

    Args:
        orderbedragen (dict): weborder → bedrag (float)
        pad_csv (str): pad naar CSV-bestand met kolommen "Weborders";"CAT"

    Returns:
        dict: categorie → totaalbedrag (float)
    """
    categorie_per_order = {}

    # 📁 CSV inlezen en mapping bouwen
    with open(pad_csv, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            weborder = row['Weborders']
            categorie = row['CAT']
            categorie_per_order[weborder] = categorie

    # 🧮 Totaalbedragen per categorie berekenen
    totaal_per_cat = defaultdict(float)
    for weborder, bedrag in orderbedragen.items():
        categorie = categorie_per_order.get(weborder)
        if categorie:
            totaal_per_cat[categorie] += bedrag
        else:
            print(f"⚠️ Weborder '{weborder}' niet gevonden in CSV.")

    return dict(totaal_per_cat)
def download_facturen_from_mail(imap_server, email_user, email_pass, SAVE_FOLDER="facturen"):
    os.makedirs(SAVE_FOLDER, exist_ok=True)

    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_user, email_pass)
    mail.select("inbox")

    status, messages = mail.search(None, 'FROM "maxim.vancompernolle@teamunited.eu"')
    mail_ids = messages[0].split()

    print(f"📨 Gevonden {len(mail_ids)} mails van maxim.vancompernolle@teamunited.eu")

    for num in mail_ids:
        status, data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        subject = decode_header(msg["Subject"])[0][0]
        print(f"\n📧 Mail onderwerp: {subject}")

        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            filename = part.get_filename()
            if filename:
                print(f"🔍 Bijlage gevonden: {filename}")
                if filename.endswith("_591050.PDF"):
                    filepath = os.path.join(SAVE_FOLDER, filename)
                    if not os.path.exists(filepath):
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        print(f"✅ Gedownload: {filename}")
                    else:
                        print(f"⏭️ Bestaat al: {filename}")

    mail.logout()    
def vind_weborders_met_posities(lines):
    resultaten = []
    for i, line in enumerate(lines):
        match = re.search(r'Website Order\s+(\d+)', line)
        if match:
            resultaten.append((str(match.group(1)), i))
    resultaten.append(("EINDE", len(lines)))
    return resultaten
def schrijf_factuurregel(
    
    factuurnummer,
    vervaldag,
    Bedrag_BTW_EXCL,
    Bedrag_BTW_incl,
    totaal_per_cat,
    pdf_naam,
    pad="facturen.csv",
  
):
    # Stap 1: check of factuurnummer al bestaat
    factuurnummers = set()
    if os.path.exists(pad):
        with open(pad, mode="r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                factuurnummers.add(row["Factuurnummer"])

    if factuurnummer in factuurnummers:
        print(f"⏭️ Factuurnummer {factuurnummer} bestaat al in {pad}, overslaan.")
        return False

    # Stap 2: schrijf nieuwe regel
    bestand_bestaat = os.path.exists(pad)
    with open(pad, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        if not bestand_bestaat:
            writer.writerow([
                "Factuurnummer", "Vervaldag", "Bedrag excl. BTW", "Bedrag incl. BTW",
                "CAT1", "CAT2", "CAT3", "CAT4", "Totaal categorieën","pdf_naam"
            ])

        writer.writerow([
            factuurnummer,
            vervaldag,
            round(Bedrag_BTW_EXCL, 2),
            round(Bedrag_BTW_incl, 2),
            round(totaal_per_cat.get("CAT1", 0.0), 2),
            round(totaal_per_cat.get("CAT2", 0.0), 2),
            round(totaal_per_cat.get("CAT3", 0.0), 2),
            round(totaal_per_cat.get("CAT4", 0.0), 2),
            round(sum(totaal_per_cat.values()), 2),
            pdf_naam
        ])

    print(f"✅ Factuur {factuurnummer} toegevoegd aan {pad}")
    return True
def verwerk_alle_facturen(pdf_map):
    pdf_bestanden = [f for f in os.listdir(pdf_map) ]

    print(f"\n📂 Verwerken van {len(pdf_bestanden)} PDF-bestanden in {pdf_map}\n")

    for pdf_naam in pdf_bestanden:
        pdf_pad = os.path.join(pdf_map, pdf_naam)
        try:
            verwerkfactuur(pdf_pad,pdf_naam)  # 👈 jouw verwerkingsfunctie
        except Exception as e:
            print(f"❌ Fout bij {pdf_naam}: {e}")    
            traceback.print_exc()
def parse_bedrag_europees(tekst):
    if not tekst or not isinstance(tekst, str):
        return None

    # Detecteer dubbele decimalen zoals '1.024.01'
    if tekst.count(",") + tekst.count(".") > 2:
        print(f"⚠️ Dubbele decimaal of fout formaat: '{tekst}'")
        return None

    # Verwijder duizendtalscheiding (punt) en vervang decimaal (komma) door punt
    clean = tekst.replace(".", "").replace(",", ".")
    
    # Laat alleen geldige float-achtige strings toe
    if not re.match(r'^\d+(\.\d{1,2})?$', clean):
        print(f"⚠️ Ongeldig bedrag: '{tekst}' ➜ '{clean}'")
        return None

    try:
        return float(clean)
    except ValueError:
        print(f"⚠️ ValueError bij bedrag: '{tekst}' ➜ '{clean}'")
        return None

download_facturen_from_mail("imap.one.com", EMAIL_USER, EMAIL_PASS, base_dir +"/pdf")
#download_facturen_debug("imap.one.com", EMAIL_USER, EMAIL_PASS, base_dir +"/pdf")
verwerk_alle_facturen(base_dir + "/pdf/")
