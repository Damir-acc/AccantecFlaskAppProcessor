from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, current_app
from flask_session import Session
import os
import extract_msg
import shutil
import re
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.graph_client import GraphClient
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from werkzeug.utils import secure_filename
import zipfile
import threading
import time
import identity.web
import requests

import application_config

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'  # Verzeichnis für hochgeladene Dateien
app.secret_key = 'supersecretkey'  # Für Flash-Nachrichten
app.config.from_object(application_config)
assert app.config["REDIRECT_PATH"] != "/", "REDIRECT_PATH must not be /"
Session(app)

app.jinja_env.globals.update(Auth=identity.web.Auth)  # Useful in template for B2C
auth = identity.web.Auth(
    session=session,
    authority=app.config["AUTHORITY"],
    client_id=app.config["CLIENT_ID"],
    client_credential=app.config["CLIENT_SECRET"],
)

@app.route(application_config.REDIRECT_PATH)
def auth_response():
    result = auth.complete_log_in(request.args)
    if "error" in result:
        return render_template("auth_error.html", result=result)
    return redirect(url_for("index"))

@app.route("/login")
def login():
    return render_template("login.html", version='1.0', **auth.log_in(
        scopes=application_config.SCOPE, # Have user consent to scopes during log-in
        redirect_uri=url_for("auth_response", _external=True, _scheme="https"), # Optional. If present, this absolute URL must match your app's redirect_uri registered in Microsoft Entra admin center
        prompt="select_account",  # Optional.
        ))

@app.route("/logout")
def logout():
    return redirect(auth.log_out(url_for("index", _external=True)))

@app.route("/")
def index():
    if not (app.config["CLIENT_ID"] and app.config["CLIENT_SECRET"]):
        return render_template('config_error.html')
    if not auth.get_user():
        return redirect(url_for("login"))
    return render_template('index.html', user=auth.get_user(), version='1.0')

@app.route("/call_downstream_api")
def call_downstream_api():
    token = auth.get_token_for_user(application_config.SCOPE)
    if "error" in token:
        return redirect(url_for("login"))
    # Use access token to call downstream api
    api_result = requests.get(
        application_config.ENDPOINT,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        timeout=30,
    ).json()
    return render_template('display.html', result=api_result)

@app.route("/processor")
def processor():
    if not auth.get_user():
        return redirect(url_for("login"))  # Weiterleitung zur Login-Seite, falls nicht eingeloggt
    return render_template('processor.html')

# Globale Variablen für Fortschritt und Status
progress = 0
status_messages = []
progress_percentage = 0  # Fortschritt in Prozent
abort_flag = False
lock = threading.Lock()  # Lock für thread-sichere Updates
emails_completed = False  # Neue Variable, um den Abschluss zu verfolgen

# Erstelle das Upload-Verzeichnis, wenn es nicht existiert
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def get_access_token():
    # Holen Sie ein Access-Token für die Microsoft Graph API
    token_response = auth.get_token_for_user(application_config.SCOPE)
    if "error" in token_response:
        raise Exception("Error getting access token: {}".format(token_response.get("error")))
    
    return token_response['access_token']


#def save_to_sharepoint_list(file_name, category, return_date, text_body, sharepoint_site_url, list_name, access_token):
def save_to_sharepoint_list(file_name, category, return_date, text_body, sharepoint_site_url, list_name, access_token):
    global lock, status_messages
   # try:
        # Get access token 
       # api_result = requests.get(
       # application_config.ENDPOINT_SHAREPOINT,
       # headers={'Authorization': 'Bearer ' + access_token['access_token']},
       # timeout=30,
       # ).json()
 
       # access_token_SH = api_result.json().get('access_token') 
      #  status_messages(f"ACCESS TOKEN SHAREPOINT: {access_token_SH}")
        # Erstellen Sie den ClientContext mit dem Access Token
       # ctx = ClientContext(sharepoint_site_url).with_access_token(access_token_SH)
        #client_credentials = ClientCredential(app.config["CLIENT_ID"],app.config["CLIENT_SECRET"])
        # Fügen Sie das Access Token direkt zu den HTTP-Headern hinzu
        #ctx.authenticate_request = lambda request: request.headers.update({'Authorization': f'Bearer {access_token}'})
        #ctx = ClientContext(sharepoint_site_url).with_access_token(access_token['access_token'])
        # Verbindungsinformationen zu SharePoint
        #ctx = ClientContext(sharepoint_site_url).with_access_token(access_token)
        #ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(user_email, user_pw))
        
        # Zugriff auf die SharePoint-Liste
       # list_object = ctx.web.lists.get_by_title(list_name)
        # Element für die SharePoint-Liste vorbereiten
      #  item_create_info = {
       #     'Title': file_name,  # Dateiname als Titel in der SharePoint-Liste
       #     'Category': category,  # Kategorie, z.B. "Out of Office", "Email-Adresse nicht gefunden"
       #     'ReturnDate': return_date.strftime('%Y-%m-%d') if return_date else 'N/A',  # Rückkehrdatum oder N/A
       #     'Email_Message': text_body,  # E-Mail-Nachricht
       # }

        # Hinzufügen des neuen Elements zur Liste
      #  list_object.add_item(item_create_info)
       # ctx.execute_query()

       # print(f"Die Datei '{file_name}' wurde erfolgreich in der SharePoint-Liste gespeichert.")
    
    #except Exception as e:
       # if hasattr(e, 'response') and e.response:
          #  status_messages.append(f"Antwort vom Server: {e.response.text}")
        # Statt den Fehler nur zu protokollieren, wird er als Exception weitergeleitet
       # raise Exception(f"Fehler beim Speichern in der SharePoint-Liste: {e}")
    

    try:
        # Token abrufen
        client_id=app.config["CLIENT_ID"]
        tenant_id=app.config["TENANT_ID"]
        client_secret=app.config["CLIENT_SECRET"]
        with lock:
            status_messages.append(f"Access Token: {access_token}")
            status_messages.append(f"CLIENT ID: {client_id}")
            status_messages.append(f"TENANT ID: {tenant_id}")
            status_messages.append(f"CLIENT ID: {client_secret}")

        
        with lock:
            status_messages.append(f"Before with Context with access token")

        #ctx = ClientContext(sharepoint_site_url).with_interactive(tenant_id, client_id)
        #with lock:
        #    status_messages.append(f"After with Context with access token, Typ von ctx: {type(ctx)}")
        #me = ctx.web.current_user.get().execute_query()
        #with lock:
        #    status_messages.append(f"With interactive, me: {me}")
        #print(me)
        #web = ctx.web.get().execute_query()
        #with lock:
        #    status_messages.append(f"With interactive, web: {web}")
        #print(web)

        ctx = ClientContext(sharepoint_site_url).with_access_token(access_token)
        target_web = ctx.web.get().execute_query()
        with lock:
            status_messages.append(f"After access token, target_web url: {target_web.url}")

        #client = GraphClient.with_client_secret(tenant_id, client_id, client_secret)
        #appli = client.applications.get_by_app_id(client_id).get().execute_query()
        #with lock:
        #    status_messages.append(f"After with Context with access token: {appli}")
        # ClientContext mit Access-Token
        #ctx = ClientContext(sharepoint_site_url).with_client_credentials(client_id, client_secret)
        #target_web = ctx.web.get().execute_query()
        #ctx = ClientContext(sharepoint_site_url).with_access_token(access_token)
        with lock:
            status_messages.append(f"After with Context with access token, Typ von ctx: {type(ctx)}")

        # Zugriff auf die SharePoint-Liste
        list_object = ctx.web.lists.get_by_title(list_name)
        with lock:
            status_messages.append(f"After access of Sharepoint-list List Object: {list_object}")
        
        # Element für die SharePoint-Liste vorbereiten
        item_create_info = {
            'Title': file_name,  # Dateiname als Titel in der SharePoint-Liste
            'Category': category,  # Kategorie, z.B. "Out of Office", "Email-Adresse nicht gefunden"
            'ReturnDate': return_date.strftime('%Y-%m-%d') if return_date else 'N/A',  # Rückkehrdatum oder N/A
            'Email_Message': text_body,  # E-Mail-Nachricht
        }
        with lock:
            status_messages.append(f"After item create info Item Create Info: {item_create_info}")


        # Neues Element hinzufügen
        list_object.add_item(item_create_info)
        with lock:
            status_messages.append(f"After list object")
        #client.applications.get_by_app_id(client_id).get().execute_query()
        ctx.execute_query()
        with lock:
            status_messages.append(f"After execute query")

        print(f"Die Datei '{file_name}' wurde erfolgreich in der SharePoint-Liste gespeichert.")
    
    except Exception as e:
        raise Exception(f"Fehler beim Speichern in der SharePoint-Liste: {e}")


# Funktion zur Extraktion des Rückkehrdatums aus einer Out-of-Office-Nachricht
def extract_return_date(message_body):
    # Aktuelles Jahr
    current_year = datetime.now().year
    current_date = datetime.now().date()

    # Mögliche Datumsformate (europäisch und amerikanisch)
    date_patterns = [
        r'\b(?:ab|von|bis|am|bis einschließlich|ab dem|on|den)\s*(\d{1,2}\.\s*(?:januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\s*\d{0,4})\b',  # Datumsformat: ab dem 23. September 2024
        r'\b(?:ab|von|bis|am|bis einschließlich|ab dem|on|den)\s*(\d{1,2}[\./]\d{1,2}[\./]?\d{0,4})\b',  # Datumsformat: 25.09 oder 25.09.2023
        r'\b(\d{1,2}\s*(?:januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\s*\d{0,4})\b',  # Datumsformat: 25. September oder 25. September 2023
        r'\b(?:vom|von|ab|bis)\s*(\d{1,2})\.\s*bis zum\s*(\d{1,2})\.\s*(\w+)\b',  # Sätze wie: vom 03. bis 26. September
        r'\b(?:on the|on|after the|until the)\s*(\d{1,2})(?:st|nd|rd|th)?\s*of\s*(\w+)',  # Format: on the 1st of October oder return to on September 23rd
        r'\b(?:on the|on|after the|until the)\s*(\w+)\s*(\d{1,2})(?:st|nd|rd|th)',  # Format: on the 1st of October oder return to on September 23rd
        r'\b(?:until)\s*(\w+)\s*(\d{1,2})(?:st|nd|rd|th)',  # Format: October 1st
        r'\b(?:until)\s*(\w+)\s*(\d{1,2})',  # Format: October 1
        r'\bam\s*\w+,\s*(\d{1,2})\.\s*([1-9]|1[0-2])\b',  # Format: am Montag, 30.9.
        r'\b(\d{1,2})\.\s*(\w+)\b',  # Format: 20. September
        r'\b(?:on the|on)\s*(\d{1,2})(?:st|nd|rd|th)\s*of\s*(\w+)\b',  # Format: 20th of September
        r'\b(?:ab dem|von|bis|am)\s*(\d{1,2}\.\s*(?:januar|februar|märz|april|mai|juni|juli|august|september|oktober|november|dezember)\b\s*\d{0,4})',  # ab 23. September oder 23. September 2024
        r'\b(\d{1,2})\.\s*(\d{1,2})\b'  # Format: 25.09
    ]

    # Zuordnung von Monatsnamen zu Zahlen
    months_dict = {
        "januar": "01", "februar": "02", "märz": "03", "april": "04", "mai": "05", "juni": "06",
        "juli": "07", "august": "08", "september": "09", "oktober": "10", "november": "11", "dezember": "12",
        "january": "01", "february": "02", "march": "03", "april": "04", "may": "05", "june": "06",
        "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12", "sept":"09",
        "jan":"01", "feb":"02", "mar":"03", "apr":"04", "jun":"06", "jul":"07", "aug":"08", "sep":"09", "oct":"10",
        "nov":"11", "dec":"12",
    }

    for pattern in date_patterns:
        matches = re.findall(pattern, message_body.lower())
        #print(matches)
        if matches:
            for match in matches:
                try:
                    # Unterscheidung zwischen verschiedenen Match-Formaten
                    if isinstance(match, tuple) and len(match) == 2:  # z.B. 20. September oder on 20th of September
                        print("Monat und Tag Format")
                        day, month = match
                        #if day in ("january","february","march","april","may","june","july","august","september","october","november","december","jan","jan.","oct","oct."):
                        #  month_temp=day
                        #  day=month
                        #  month=month_temp
                        #  month = month.rstrip(".")
                        if day in months_dict:  # Wenn der Tag eigentlich der Monat ist (Fehler)
                            day, month = month, day
                        month = months_dict.get(month.lower(), month)  # Monat als Zahl umwandeln
                        day = day.rstrip("stndrdth")  # Entfernen der englischen Suffixe wie 1st, 2nd
                        date_str = f"{day}.{month}.{current_year}"

                    elif isinstance(match, tuple) and len(match) == 3:  # vom 20. bis 26. September
                        print("Von-bis Format")
                        _, day_end, month = match
                        month = months_dict.get(month.lower(), month)  # Monat als Zahl umwandeln
                        date_str = f"{day_end}.{month}.{current_year}"

                    else:
                        print("Normales Format")
                        date_str = match
                        date_str = date_str.strip()  # Eventuelle zusätzliche Leerzeichen entfernen


                    # Falls das Datum den Monat ausgeschrieben enthält, umformatieren
                    if any(month in date_str.lower() for month in months_dict.keys()):
                        for month_name, month_num in months_dict.items():
                            date_str = date_str.replace(month_name, month_num)  


                    # Sicherstellen, dass die Punkte richtig gesetzt sind
                    date_str = re.sub(r'(\d{1,2})\s*\.(\d{1,2})', r'\1.\2', date_str)  # Sicherstellen, dass nach dem Tag ein Punkt steht
                    date_str = re.sub(r'(\d{1,2})\s+(?=\d{4})', r'\1.', date_str)  # Punkt setzen, wenn Jahr folgt
                    date_str = re.sub(r'\s*\.\s*', '.', date_str)  # Leerzeichen um Punkte entfernen
                    date_str = re.sub(r'\.{2,}', '.', date_str)  # Doppelte Punkte entfernen
                    date_str = date_str.strip('.')

                    # Prüfen, ob das Jahr fehlt und es hinzufügen
                    parts = date_str.split('.')
                    print(parts)
                    if len(parts[0])==1:
                       parts[0] = parts[0].zfill(2)
                    if len(parts[1])==1:
                       parts[1] = parts[1].zfill(2)
                    if len(parts) == 2:  # Wenn nur Tag und Monat vorhanden sind
                        date_str = f"{date_str}.{current_year}"
                    elif len(parts) == 3 and len(parts[2]) == 2:  # zweistelliges Jahr (z.B. 24)
                        year = int(parts[2])
                        year_full = 2000 + year if year <= current_year % 100 else 1900 + year  # Umwandlung in vierstelliges Jahr
                        date_str = f"{parts[0]}.{parts[1]}.{year_full}"
                    elif len(parts) == 3 and len(parts[2]) == 4:  # vierstelliges Jahr bereits vorhanden
                        date_str = f"{parts[0]}.{parts[1]}.{parts[2]}"  # Keine Änderung nötig

                    print(date_str)
                    # Validierung des Datumsstrings
                    if not re.match(r'^\d{2}\.\d{2}\.\d{4}$', date_str):
                        print("Kein Datum")
                        return None  # Rückgabe None, wenn das Datum nicht das gewünschte Format hat
                    # Parsing des Datums
                    return_date = datetime.strptime(date_str, "%d.%m.%Y").date()

                except Exception as e:
                    print(f"Fehler bei der Verarbeitung des Datums: {e}")
                return return_date
    
    return None

# Funktion, um die Kategorie basierend auf dem Textinhalt und Betreff zu bestimmen
def categorize_message(subject, message_body):
    # Konvertiere den Betreff und Nachrichtentext in Kleinbuchstaben für die Suche
    lower_subject = subject.lower()
    lower_body = message_body.lower()

    pattern = re.compile("nicht mehr für .+ tätig")
    #print(lower_subject)
    #print(lower_body)
    # Überprüfung auf bestimmte Bedingungen im Betreff oder Nachrichtenkörper
    if "out of office" in lower_body or "abwesend" in lower_body or "not available" in lower_body or "wieder persönlich für sie da" in lower_body or "im büro erreichbar" in lower_body or "erreichen mich wieder" in lower_body or "on leave" in lower_body or "im urlaub" in lower_body or "elternzeit" in lower_body or "nicht im dienst" in lower_body or "wieder im hause" in lower_body or "out of the office" in lower_body or "dienstreise" in lower_body or "geschäftsreise" in lower_body or "abwesenheit" in lower_body or "on vacation" in lower_body or "wieder erreichbar" in lower_body or "außer haus" in lower_body or "nicht im haus" in lower_body or "nicht erreichbar" in lower_body or "nicht im büro" in lower_body or "out of office" in lower_subject or "abwesenheit" in lower_subject:
        return_date = extract_return_date(lower_body)
        print(return_date)
        return "Out of Office"
    elif "email address does not exist" in lower_body or "existiert nicht mehr" in lower_body or "no longer with" in lower_body or "no longer employed" in lower_body or "nicht mehr tätig" in lower_body or "retirement" in lower_body or "nicht mehr beschäftigt" in lower_body or "unzustellbar" in lower_subject or "email address does not exist" in lower_subject or "undelivered mail" in lower_subject or pattern.search(lower_body):
        return "Inaktive E-Mail"
    elif "unsubscribe" in lower_body:
        return "Abbestellen"
    elif "aw" in lower_subject or "re" in lower_subject or "sehr geehrter herr" in lower_body or "sehr geehrte frau" in lower_body or "hallo frau" in lower_body or "hallo herr" in lower_body or "guten tag frau" in lower_body or "guten tag herr" in lower_body or "liebe frau" in lower_body or "lieber herr" in lower_body or "guten morgen frau" in lower_body or "guten morgen herr" in lower_body:
        return "Rückmeldung"
    else:
        return "Unkategorisiert"

def process_and_copy_messages(file_path, sharepoint_site_url, list_name, access_token):
    global progress, progress_percentage, lock, abort_flag, status_messages
    if file_path.endswith(".msg"):
        msg = extract_msg.Message(file_path)
        msg_body = msg.body
        msg_subject = msg.subject
        category = categorize_message(msg_subject, msg_body)
        return_date = None
        if category == "Out of Office":
            return_date = extract_return_date(msg_body)

        # Hier wird der Zielordner festgelegt (kann angepasst werden)
        # In diesem Fall speichern wir die Dateien nicht lokal, sondern nur in SharePoint
        try:
            save_to_sharepoint_list(os.path.basename(file_path), category, return_date, msg_body, sharepoint_site_url, list_name, access_token)
        except Exception as e:
            with lock:
                status_messages.append(f"Fehler beim Hochladen der E-Mail: {e}")
            # Setze das Abbruchflag und beende den Thread
            with lock:
                abort_flag = True
            return

        # Thread-sichere Fortschrittsaktualisierung
        with lock:
            progress += 1

def email_processing_thread(file_paths, sharepoint_site_url, list_name, access_token):
    global progress, progress_percentage, lock, abort_flag, emails_completed, status_messages
    total_files = len(file_paths)

    for file_path in file_paths:
        # Abbruchprüfung
        if abort_flag:
            with lock:
                status_messages.append("Hochladen wurde abgebrochen.")
            break

        process_and_copy_messages(file_path, sharepoint_site_url, list_name, access_token)
        
        # Thread-sichere Berechnung des Fortschritts
        with lock:
            status_messages.append(f"E-Mail {progress}/{total_files} hochgeladen.")
            progress_percentage = int((progress / total_files) * 100)

    # Kopieren abgeschlossen oder abgebrochen
    with lock:
        emails_completed = True
    

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    global progress,progress_percentage, abort_flag, emails_completed, status_messages
    # Fortschritt und Statusmeldungen beim Neuladen der Seite zurücksetzen
    if request.method == 'GET':
        with lock:  # Thread-Safe Zurücksetzen
            progress = 0
            progress_percentage = 0
            status_messages= []
            abort_flag = False  # Reset des Abbruch-Flags
            emails_completed = False
    if request.method == 'POST':
        # Setze abort_flag zurück, bevor ein neuer Upload-Prozess gestartet wird
        with lock:
            abort_flag = False  # Reset des Abbruch-Flags bei POST-Start       
        # Überprüfe, ob die Datei im Request vorhanden ist
        if 'file' not in request.files:
            status_messages.append('Keine Datei ausgewählt.')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            status_messages.append('Keine Datei ausgewählt.')
            return redirect(request.url)

        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Überprüfe, ob die Datei eine ZIP-Datei ist
            if filename.endswith('.zip'):
                # Entpacken der ZIP-Datei
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(app.config['UPLOAD_FOLDER'])

                # Verarbeitung aller .msg-Dateien im entpackten Verzeichnis
                file_paths = []
                for root, dirs, files in os.walk(app.config['UPLOAD_FOLDER']):
                    for f in files:
                        if f.endswith('.msg'):
                            msg_file_path = os.path.join(root, f)
                            file_paths.append(msg_file_path)

                # Holen der SharePoint-Daten aus dem Formular
                sharepoint_site_url = request.form.get('sharepoint_url')
                list_name = request.form.get('list_name')
                user_email = request.form.get('user_email')
                user_pw = request.form.get('user_pw')

                if not all([sharepoint_site_url, list_name, user_email, user_pw]):
                    status_messages.append('Bitte fülle alle Felder aus.')
                    return redirect(request.url)
                
                # Initialisiere den Fortschritt
                #progress = 0

                #access_token = auth.get_token_for_user(application_config.SCOPE)
                access_token=get_access_token()

                # Starte den E-Mail-Verarbeitungs-Thread
                threading.Thread(target=email_processing_thread, args=(file_paths, sharepoint_site_url, list_name, access_token)).start()
                #threading.Thread(target=email_processing_thread, args=(file_paths, sharepoint_site_url, list_name)).start()
                status_messages.append('Dateien werden verarbeitet.')

                # Löschen der ZIP-Datei nach der Verarbeitung
                os.remove(file_path)

            else:
                status_messages.append('Bitte laden Sie eine ZIP-Datei mit .msg-Dateien hoch.')
                return redirect(request.url)

            return redirect(url_for('upload_files'))

    return render_template('processor.html')

@app.route('/api/abort', methods=['POST'])
def abort():
    global abort_flag
    with lock:
        abort_flag = True  # Setze das Abbruch-Flag
        status_messages.append("Abbruchvorgang wurde eingeleitet.")
    return jsonify({"message": "Abbruchvorgang wurde eingeleitet."}), 200

@app.route('/api/status', methods=['GET'])
def get_status():
    with lock:  # Thread-Safe Status auslesen
        return jsonify(status_messages), 200

@app.route('/api/progress', methods=['GET'])
def get_progress():
    global progress_percentage
    with lock:  # Thread-Safe Fortschritt auslesen
        return jsonify({"progress": progress_percentage}), 200
    
@app.route('/api/complete', methods=['GET'])
def check_complete():
    global emails_completed
    with lock:
        return jsonify({"completed": emails_completed}), 200

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
