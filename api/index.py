import os
import json
import urllib.parse
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response
import csv
import io
import traceback

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'templates')
app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config['TEMPLATES_AUTO_RELOAD'] = True

SHEET_ID = '1ATeUpMRTcKLA6PJLONjbOI1eMMmZV5ifCpsmhJoqQu4'
SEM_NAMES = ['Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']
# Columnas en hoja Asistencia (0-indexed): 0=Apellidos,1=Nombre,2=Email,3=Inst,4=Pais,5=Timestamp,6=Recep,7=Sem1,8=Sem2,9=Sem3,10=Sem4
SEM_COL_INDEX = {'Seminario 1': 7, 'Seminario 2': 8, 'Seminario 3': 9, 'Seminario 4': 10}

def get_gc():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if creds_json:
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        local_creds = os.path.join(os.path.dirname(__file__), '..', 'credentials.json')
        if os.path.exists(local_creds):
            creds = Credentials.from_service_account_file(local_creds, scopes=scopes)
        else:
            return None
    return gspread.authorize(creds)

def get_sheet():
    gc = get_gc()
    if not gc:
        return None
    return gc.open_by_key(SHEET_ID)

def normalize_key(apellidos, nombre):
    return (apellidos.strip().lower(), nombre.strip().lower())

# =========================================================
# CARGA DE ASISTENTES (cache en cold start para búsqueda)
# =========================================================
def load_attendees():
    sh = get_sheet()
    if not sh:
        return []
    try:
        ws_ai = sh.worksheet('Asistentes_imprimir')
        rows_ai = ws_ai.get_all_values()[1:]
        ws_main = sh.worksheet('Asistentes')
        rows_main = ws_main.get_all_values()[1:]

        email_lookup = {}
        for row in rows_main:
            if len(row) >= 3:
                key = normalize_key(row[0], row[1])
                email = row[2].strip()
                if email and email.lower() != 'nan':
                    email_lookup[key] = email

        sem_membership = {}
        for sem in SEM_NAMES:
            try:
                ws_sem = sh.worksheet(sem)
                vals = ws_sem.get_all_values()[1:]
                for row in vals:
                    if len(row) >= 2:
                        key = normalize_key(row[0], row[1])
                        if key not in sem_membership:
                            sem_membership[key] = []
                        if sem not in sem_membership[key]:
                            sem_membership[key].append(sem)
                    if len(row) >= 3:
                        email = row[2].strip()
                        if email and email.lower() != 'nan':
                            email_lookup[key] = email
            except:
                pass

        attendees = []
        for row in rows_ai:
            if len(row) < 3:
                continue
            apellidos = row[1].strip() if len(row) > 1 else ''
            nombre = row[2].strip() if len(row) > 2 else ''
            if not apellidos or not nombre:
                continue
            key = normalize_key(apellidos, nombre)
            email = email_lookup.get(key, '')
            seminarios = sem_membership.get(key, [])
            institucion = row[3].strip() if len(row) > 3 else ''
            pais = row[4].strip() if len(row) > 4 else ''
            attendees.append({
                'apellidos': apellidos, 'nombre': nombre, 'email': email,
                'institucion': institucion, 'pais': pais,
                'seminarios': seminarios, 'total_seminarios': len(seminarios),
                'key_id': email if email else f"{apellidos.lower()}_{nombre.lower()}",
                'tipo_inscripcion': 'Solo Congreso' if len(seminarios) == 0 else f'Congreso + {len(seminarios)} Seminario(s)'
            })
        return attendees
    except Exception as e:
        print(f'Error cargando asistentes: {e}')
        traceback.print_exc()
        return []

# =========================================================
# OPERACIONES DIRECTAS EN HOJA "Asistencia" (siempre Sheets)
# =========================================================
ASIST_HEADERS = ['Apellidos', 'Nombre', 'Email', 'Institución', 'País', 'Timestamp', 'Recepcionista', 'Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']

def get_asistencia_ws(sh):
    try:
        return sh.worksheet('Asistencia')
    except:
        ws = sh.add_worksheet('Asistencia', 1000, 12)
        ws.append_row(ASIST_HEADERS)
        return ws

def normalize_for_compare(text):
    """Normaliza texto para comparación, removiendo acentos"""
    import unicodedata
    text = text.lower().strip()
    # Remover acentos
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

def find_in_asistencia(ws, key_id):
    all_vals = ws.get_all_values()
    key_id_normalized = normalize_for_compare(key_id)
    for i, row in enumerate(all_vals):
        if i == 0:
            continue
        if len(row) >= 3:
            row_email = row[2].strip()
            row_key = f"{row[0].strip()}_{row[1].strip()}"
            row_key_normalized = normalize_for_compare(row_key)
            
            # Comparar por email exacto o por key normalizado
            if (row_email and row_email == key_id) or row_key_normalized == key_id_normalized:
                return row, i + 1
    return None, None

def read_status_from_sheets(key_id):
    sh = get_sheet()
    if not sh:
        return {'asistio': False, 'timestamp': '', 'seminarios': {}}
    try:
        ws = get_asistencia_ws(sh)
        row, _ = find_in_asistencia(ws, key_id)
        if not row:
            return {'asistio': False, 'timestamp': '', 'seminarios': {}}
        seminarios = {}
        for sem, idx in SEM_COL_INDEX.items():
            if len(row) > idx and row[idx]:
                seminarios[sem] = row[idx]
        return {
            'asistio': bool(len(row) > 5 and row[5]),
            'timestamp': row[5] if len(row) > 5 else '',
            'seminarios': seminarios
        }
    except Exception as e:
        print(f'Error leyendo status: {e}')
        return {'asistio': False, 'timestamp': '', 'seminarios': {}}

def write_general_attendance(attendee, recepcionista):
    sh = get_sheet()
    if not sh:
        return {'status': 'error', 'message': 'Sin conexión a Sheets'}
    try:
        ws = get_asistencia_ws(sh)
        row, row_num = find_in_asistencia(ws, attendee['key_id'])
        if row and len(row) > 5 and row[5]:
            return {'status': 'already', 'timestamp': row[5]}
        timestamp = datetime.now().isoformat()
        new_row = [
            attendee['apellidos'], attendee['nombre'], attendee.get('email', ''),
            attendee.get('institucion', ''), attendee.get('pais', ''),
            timestamp, recepcionista, '', '', '', ''
        ]
        if row_num:
            for sem, idx in SEM_COL_INDEX.items():
                if len(row) > idx:
                    new_row[idx] = row[idx]
            ws.update(f'A{row_num}:K{row_num}', [new_row])
        else:
            ws.append_row(new_row)
        return {'status': 'ok', 'timestamp': timestamp}
    except Exception as e:
        print(f'Error marcando asistencia: {e}')
        traceback.print_exc()
        return {'status': 'error', 'message': str(e)}

def write_seminar_attendance(attendee, seminario, recepcionista):
    sh = get_sheet()
    if not sh:
        return {'status': 'error', 'message': 'Sin conexión a Sheets'}
    try:
        ws = get_asistencia_ws(sh)
        row, row_num = find_in_asistencia(ws, attendee['key_id'])
        sem_idx = SEM_COL_INDEX.get(seminario)
        if sem_idx is None:
            return {'status': 'error', 'message': 'Seminario no válido'}
        timestamp = datetime.now().isoformat()
        if row_num:
            if len(row) > sem_idx and row[sem_idx]:
                return {'status': 'already'}
            col_letter = chr(65 + sem_idx)
            ws.update(f'{col_letter}{row_num}', [[timestamp]])
        else:
            new_row = [
                attendee['apellidos'], attendee['nombre'], attendee.get('email', ''),
                attendee.get('institucion', ''), attendee.get('pais', ''),
                '', recepcionista,
                timestamp if seminario == 'Seminario 1' else '',
                timestamp if seminario == 'Seminario 2' else '',
                timestamp if seminario == 'Seminario 3' else '',
                timestamp if seminario == 'Seminario 4' else '',
            ]
            ws.append_row(new_row)
        return {'status': 'ok', 'timestamp': timestamp}
    except Exception as e:
        print(f'Error marcando seminario: {e}')
        traceback.print_exc()
        return {'status': 'error', 'message': str(e)}

def read_all_attendance():
    sh = get_sheet()
    if not sh:
        return {}
    try:
        ws = sh.worksheet('Asistencia')
        all_vals = ws.get_all_values()
        att = {}
        for row in all_vals[1:]:
            if len(row) >= 3 and (row[0] or row[1]):
                email = row[2].strip() if len(row) > 2 else ''
                key = email if email else f"{row[0].strip().lower()}_{row[1].strip().lower()}"
                seminarios = {}
                for sem, idx in SEM_COL_INDEX.items():
                    if len(row) > idx and row[idx]:
                        seminarios[sem] = row[idx]
                att[key] = {'timestamp': row[5] if len(row) > 5 else '', 'seminarios': seminarios}
        return att
    except:
        return {}

# =========================================================
# Cold start
# =========================================================
attendees = load_attendees()

# =========================================================
# RUTAS
# =========================================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/search')
def search():
    q = request.args.get('q', '').strip().lower()
    if not q or len(q) < 2:
        return jsonify([])
    results = []
    for a in attendees:
        full = f"{a['nombre']} {a['apellidos']}".lower()
        rev = f"{a['apellidos']} {a['nombre']}".lower()
        email = a['email'].lower()
        if q in full or q in rev or q in a['nombre'].lower() or q in a['apellidos'].lower() or (email and q in email):
            results.append(a)
            if len(results) >= 30:
                break
    return jsonify(results)

@app.route('/api/status/<path:key_id>')
def api_status(key_id):
    decoded_key = urllib.parse.unquote(key_id)
    return jsonify(read_status_from_sheets(decoded_key))

@app.route('/api/mark', methods=['POST'])
def api_mark():
    data = request.json
    key_id = data.get('key_id', '')
    recepcionista = data.get('recepcionista', '')
    if not key_id:
        return jsonify({'error': 'ID requerido'}), 400
    # Normalizar para comparación
    key_id_normalized = key_id.strip()
    attendee = next((a for a in attendees if a['key_id'] == key_id_normalized), None)
    if not attendee:
        return jsonify({'error': 'Asistente no encontrado'}), 404
    return jsonify(write_general_attendance(attendee, recepcionista))

@app.route('/api/mark-seminar', methods=['POST'])
def api_mark_seminar():
    data = request.json
    key_id = data.get('key_id', '')
    seminario = data.get('seminario', '')
    recepcionista = data.get('recepcionista', '')
    if not key_id or not seminario:
        return jsonify({'error': 'ID y seminario requeridos'}), 400
    attendee = next((a for a in attendees if a['key_id'] == key_id), None)
    if not attendee:
        return jsonify({'error': 'Asistente no encontrado'}), 404
    return jsonify(write_seminar_attendance(attendee, seminario, recepcionista))

@app.route('/api/stats')
def api_stats():
    att = read_all_attendance()
    sem_stats = {}
    for sem in SEM_NAMES:
        inscritos = sum(1 for a in attendees if sem in a['seminarios'])
        asistieron = sum(1 for a in attendees if sem in a['seminarios'] and att.get(a['key_id'], {}).get('seminarios', {}).get(sem))
        sem_stats[sem] = {'inscritos': inscritos, 'asistieron': asistieron}
    marcados = sum(1 for v in att.values() if v.get('timestamp'))
    return jsonify({
        'total': len(attendees), 'asistieron': marcados, 'pendientes': len(attendees) - marcados,
        'solo_congreso': sum(1 for a in attendees if a['total_seminarios'] == 0),
        'con_seminarios': sum(1 for a in attendees if a['total_seminarios'] > 0),
        'seminarios': sem_stats
    })

@app.route('/api/reload', methods=['POST'])
def api_reload():
    global attendees
    attendees = load_attendees()
    return jsonify({'status': 'ok', 'total': len(attendees)})

@app.route('/report')
def report():
    return render_template('report.html')

@app.route('/api/report-data')
def report_data():
    att = read_all_attendance()
    rows = []
    for a in attendees:
        rec = att.get(a['key_id'], {})
        sem_a = rec.get('seminarios', {})
        rows.append({
            'nombre_completo': f"{a['nombre']} {a['apellidos']}",
            'apellidos': a['apellidos'], 'nombre': a['nombre'], 'email': a['email'],
            'institucion': a['institucion'], 'pais': a['pais'], 'tipo': a['tipo_inscripcion'],
            'asistio_general': bool(rec.get('timestamp')),
            'seminarios_inscritos': a['seminarios'], 'total_inscritos': a['total_seminarios'],
            'seminarios_asistidos': sem_a, 'total_asistidos': len(sem_a),
            'Seminario 1': 'Seminario 1' in sem_a, 'Seminario 2': 'Seminario 2' in sem_a,
            'Seminario 3': 'Seminario 3' in sem_a, 'Seminario 4': 'Seminario 4' in sem_a,
        })
    sem_stats = {}
    for sem in SEM_NAMES:
        ins = sum(1 for a in attendees if sem in a['seminarios'])
        asi = sum(1 for r in rows if r[sem])
        sem_stats[sem] = {'inscritos': ins, 'asistieron': asi, 'porcentaje': round(asi/ins*100,1) if ins > 0 else 0}
    tg = sum(1 for r in rows if r['asistio_general'])
    return jsonify({'total': len(rows), 'asistieron_general': tg, 'pendientes_general': len(rows)-tg,
        'pct_general': round(tg/len(rows)*100,1) if rows else 0, 'seminarios': sem_stats, 'rows': rows})

@app.route('/api/report-csv')
def report_csv():
    att = read_all_attendance()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Apellidos','Nombre','Email','Institucion','Pais','Tipo','Asistio General',
                     'Seminarios Inscritos','Seminarios Asistidos','Sem 1','Sem 2','Sem 3','Sem 4'])
    for a in attendees:
        rec = att.get(a['key_id'], {})
        sa = rec.get('seminarios', {})
        writer.writerow([a['apellidos'],a['nombre'],a['email'],a['institucion'],a['pais'],
            a['tipo_inscripcion'], 'Si' if rec.get('timestamp') else 'No',
            ', '.join(a['seminarios']), ', '.join(sa.keys()),
            'Si' if 'Seminario 1' in sa else 'No', 'Si' if 'Seminario 2' in sa else 'No',
            'Si' if 'Seminario 3' in sa else 'No', 'Si' if 'Seminario 4' in sa else 'No'])
    return Response(output.getvalue(), mimetype='text/csv',
                    headers={'Content-Disposition': 'attachment; filename=reporte_cruc_2026.csv'})

@app.route('/api/register-last-minute', methods=['POST'])
def register_last_minute():
    global attendees
    data = request.json
    apellidos = data.get('apellidos', '').strip()
    nombre = data.get('nombre', '').strip()
    email = data.get('email', '').strip()
    institucion = data.get('institucion', '').strip()
    pais = data.get('pais', '').strip()
    seminarios = data.get('seminarios', [])
    if not apellidos or not nombre:
        return jsonify({'error': 'Apellidos y nombre obligatorios'}), 400
    key_id = email if email else f"{apellidos.lower()}_{nombre.lower()}"
    if any(a['key_id'] == key_id for a in attendees):
        return jsonify({'error': 'Ya existe'}), 400
    new_a = {
        'apellidos': apellidos, 'nombre': nombre, 'email': email,
        'institucion': institucion, 'pais': pais,
        'seminarios': seminarios, 'total_seminarios': len(seminarios),
        'key_id': key_id,
        'tipo_inscripcion': 'Solo Congreso' if len(seminarios) == 0 else f'Congreso + {len(seminarios)} Seminario(s)'
    }
    sh = get_sheet()
    gsheet_ok = False
    if sh:
        try:
            ws_ai = sh.worksheet('Asistentes_imprimir')
            next_num = len(ws_ai.get_all_values())
            ws_ai.append_row([next_num, apellidos, nombre, institucion, pais])
            for sem in seminarios:
                try:
                    ws_sem = sh.worksheet(sem)
                    ws_sem.append_row([apellidos, nombre, email, institucion, pais])
                except:
                    pass
            gsheet_ok = True
        except Exception as e:
            print(f'Error escribiendo nuevo asistente: {e}')
    attendees.append(new_a)
    return jsonify({'status': 'ok', 'gsheet_ok': gsheet_ok, 'total': len(attendees)})

@app.route('/api/diagnostic')
def diagnostic():
    gc = get_gc()
    return jsonify({'gc_ok': gc is not None, 'total_asistentes': len(attendees)})

# Vercel: expone 'app' como handler WSGI
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f'Servidor iniciado en http://localhost:{port}')
    print(f'Asistentes cargados: {len(attendees)}')
    app.run(debug=False, host='0.0.0.0', port=port)
