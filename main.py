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
import smtplib
from email.message import EmailMessage
import mimetypes


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
    def verplaats_mail_naar_behandeld(mail, mail_id, foldernaam="INBOX.behandeldefacturenBrandsfit"):
        """
        Verplaatst een mail met mail_id naar de opgegeven IMAP-folder.
        Maakt de folder aan als die nog niet bestaat.
        Debug: print IMAP responses.
        """
        # Controleer of de folder bestaat, anders aanmaken
        typ, folders = mail.list()
        print("IMAP folders:", folders)
        folder_bestaat = any(foldernaam in f.decode(errors="ignore") for f in folders if isinstance(f, bytes))
        if not folder_bestaat:
            print(f"Folder '{foldernaam}' bestaat niet, aanmaken...")
            resp = mail.create(foldernaam)
            print("Create response:", resp)
        # Verplaats de mail
        print(f"Mail {mail_id} kopiëren naar {foldernaam}...")
        result = mail.copy(mail_id, foldernaam)
        print("Copy response:", result)
        if result[0] == 'OK':
            # Markeer als verwijderd in INBOX
            store_result = mail.store(mail_id, '+FLAGS', '\\Deleted')
            print("Store (delete flag) response:", store_result)
            print(f"📥 Mail {mail_id.decode() if isinstance(mail_id, bytes) else mail_id} verplaatst naar '{foldernaam}'")
        else:
            print(f"⚠️ Kon mail {mail_id.decode() if isinstance(mail_id, bytes) else mail_id} niet verplaatsen naar '{foldernaam}'")
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_user, email_pass)
    mail.select("inbox")

    status, messages = mail.search(None, 'FROM "maxim.vancompernolle@teamunited.eu"')
    mail_ids = messages[0].split()

    print(f"📨 Gevonden {len(mail_ids)} mails van maxim.vancompernolle@teamunited.eu")

    gedownloade_pdfs = []

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
                        gedownloade_pdfs.append(filename) 
                        print(f"✅ Gedownload: {filename}")
                    else:
                        print(f"⏭️ Bestaat al: {filename}")
                    verplaats_mail_naar_behandeld(mail, num)
    mail.expunge()  # Verwijder gemarkeerde mails definitief uit INBOX
    mail.logout()

    # Overzicht van alle gedownloade PDF-bestanden
    return gedownloade_pdfs
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
    nieuwe_facturen = download_facturen_from_mail("imap.one.com", EMAIL_USER, EMAIL_PASS, base_dir +"/pdf")
    pdf_bestanden = [f for f in os.listdir(pdf_map)]

    print(f"\n📂 Verwerken van {len(pdf_bestanden)} PDF-bestanden in {pdf_map}\n")

    for pdf_naam in pdf_bestanden:
        pdf_pad = os.path.join(pdf_map, pdf_naam)
        try:
            verwerkfactuur(pdf_pad, pdf_naam)
        except Exception as e:
            print(f"❌ Fout bij {pdf_naam}: {e}")
            traceback.print_exc()

    # Verstuur alleen als er nieuwe facturen zijn
    if nieuwe_facturen:
        verstuur_nieuwe_facturen_mail(
            ontvanger="jasper.claeys@gmail.com",
            csv_pad="facturen.csv",
            pdf_map=base_dir + "/pdf/",
            pdf_namen=nieuwe_facturen,
            smtp_server="send.one.com",
            smtp_port=587,
            smtp_user=EMAIL_USER,
            smtp_pass=EMAIL_PASS
        )
    else:
        print("Geen nieuwe facturen om te mailen.")
def verstuur_nieuwe_facturen_mail(
    ontvanger="jasper.claeys@gmail.com",
    csv_pad="facturen.csv",
    pdf_map="pdf/",
    pdf_namen=None,
    smtp_server="send.one.com",
    smtp_port=587,
    smtp_user=EMAIL_USER,
    smtp_pass=EMAIL_PASS
):
    """
    Verstuur een mail met de nieuwe facturen (CSV en PDF's) als bijlage.
    Alleen de nieuwe PDF's worden toegevoegd en vermeld in de body.
    """
    if not pdf_namen:
        print("⚠️ Geen nieuwe PDF-bestanden om te versturen.")
        return

    msg = EmailMessage()
    msg["Subject"] = "Nieuwe facturen ter verwerking voor betaling"
    msg["From"] = smtp_user
    msg["To"] = ontvanger

    # Body met overzicht van de nieuwe facturen
    body = (
        "Beste,\n\n"
        "In de bijlage vind je het overzicht van de nieuwe facturen (CSV) en de volgende PDF-bestanden:\n\n"
        + "\n".join(f"- {naam}" for naam in pdf_namen) +
        "\n\nGelieve deze te verwerken voor betaling.\n\nMet vriendelijke groeten,\nAutomatisch script"
    )
    msg.set_content(body)

    # Voeg CSV toe als bijlage
    with open(csv_pad, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="text",
            subtype="csv",
            filename=os.path.basename(csv_pad)
        )

    # Voeg alleen de nieuwe PDF's toe als bijlage
    for pdf_naam in pdf_namen:
        pdf_pad = os.path.join(pdf_map, pdf_naam)
        if os.path.exists(pdf_pad):
            ctype, encoding = mimetypes.guess_type(pdf_pad)
            maintype, subtype = ctype.split("/", 1) if ctype else ("application", "pdf")
            with open(pdf_pad, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=pdf_naam
                )

    # Verstuur de mail via SMTP
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        print(f"✅ Mail met nieuwe facturen verstuurd naar {ontvanger}")

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


def lees_orderbevestigingen_en_append_orders(imap_server, email_user, email_pass, orders_csv="orders.csv"):
    """
    Leest mails met onderwerp 'Orderbevestiging' van info@teamunited.eu,
    zoekt het referentienummer en voegt deze toe aan orders.csv na prompt voor categorie.
    Voegt alleen toe als het referentienummer nog niet in de CSV staat.
    """
    # Bestaande referenties inlezen
    bestaande_refs = set()
    try:
        with open(orders_csv, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                ref = row.get("Weborders") or row.get('"Weborders"')
                if ref:
                    bestaande_refs.add(ref.strip().strip('"'))
    except FileNotFoundError:
        pass  # CSV bestaat nog niet

    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_user, email_pass)
    mail.select("inbox")

    status, messages = mail.search(None, '(FROM "info@teamunited.eu" SUBJECT "Orderbevestiging")')
    mail_ids = messages[0].split()

    nieuwe_orders = []

    for num in mail_ids:
        status, data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        subject = email.header.decode_header(msg["Subject"])[0][0]
        if isinstance(subject, bytes):
            subject = subject.decode()
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8")
                    break
        else:
            body = msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8")

        # Zoek referentienummer
        match = re.search(r"referentie\s+(\d+)", body)
        if match:
            referentie = match.group(1)
            if referentie in bestaande_refs:
                print(f"⏭️ Referentie {referentie} bestaat al in {orders_csv}, wordt overgeslagen.")
                continue
            print(f"Gevonden referentie: {referentie}")
            # Prompt voor categorie
            while True:
                print   ("Cat 1	Kledij die hoort bij wat leden in hun lidgeld inbegrepen krijgen (de 2 jaarlijkse training en zijn nabestellingen)   [speciale gevallen mbt btw recuperatie vs het feit dat we op lidgeld geen Btw moeten betalen] ")
                print   ("CAT 2	Kledij die we als organisatie ter beschikking stellen van medewekers ( jassen voor trainers, afgevaardigden enz)  of van ploegen ( de gesponsorde outfits)  [Investeringen]")
                print   ("CAT 3	 Kledij die we als verbruiksgoed beschouwen ( broekjes , kousen ) [verbruiken, kosten]")
                print   ("CAT 4	 Kledij die we doorverkopen [commerce]")
                cat = input(f"Geef categorie voor order {referentie} (CAT1, CAT2, CAT3, CAT4): ").strip().upper()
                if cat in {"CAT1", "CAT2", "CAT3", "CAT4"}:
                    break
                print("Ongeldige categorie, probeer opnieuw.")
            nieuwe_orders.append((referentie, cat))
        else:
            print("⚠️ Geen referentie gevonden in mail.")

    # Append nieuwe orders aan CSV
    if nieuwe_orders:
        with open(orders_csv, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=';')
            for ref, cat in nieuwe_orders:
                writer.writerow([ref, cat])
        print(f"{len(nieuwe_orders)} nieuwe orders toegevoegd aan {orders_csv}")
    else:
        print("Geen nieuwe orders gevonden.")

    mail.logout()



#download_facturen_debug("imap.one.com", EMAIL_USER, EMAIL_PASS, base_dir +"/pdf")
lees_orderbevestigingen_en_append_orders("imap.one.com", EMAIL_USER, EMAIL_PASS, "orders.csv")
#verwerk_alle_facturen(base_dir + "/pdf/")