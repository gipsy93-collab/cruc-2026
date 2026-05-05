import os
import sys
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response
import csv
import io
import traceback

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'templates')
app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config['TEMPLATES_AUTO_RELOAD'] = True

SHEET_ID = '1ATeUpMRTcKLA6PJLONjbOI1eMMmZV5ifCpsmhJoqQu4'

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

def normalize_key(apellidos, nombre):
    return (apellidos.strip().lower(), nombre.strip().lower())

def load_data():
    gc = get_gc()
    if not gc:
        return [], {}
    try:
        sh = gc.open_by_key(SHEET_ID)
        ws_ai = sh.worksheet('Asistentes_imprimir')
        rows_ai = ws_ai.get_all_values()[1:]
        ws_main = sh.worksheet('Asistentes')
        rows_main = ws_main.get_all_values()[1:]

        sem_names = ['Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']
        sem_data = {}
        for sem in sem_names:
            try:
                ws_sem = sh.worksheet(sem)
                vals = ws_sem.get_all_values()[1:]
                sem_data[sem] = vals
            except:
                sem_data[sem] = []

        email_lookup = {}
        for row in rows_main:
            if len(row) >= 3:
                key = normalize_key(row[0], row[1])
                email = row[2].strip()
                if email and email.lower() != 'nan':
                    email_lookup[key] = email
        for sem in sem_names:
            for row in sem_data[sem]:
                if len(row) >= 3:
                    key = normalize_key(row[0], row[1])
                    email = row[2].strip()
                    if email and email.lower() != 'nan':
                        email_lookup[key] = email

        sem_membership = {}
        for sem_name in sem_names:
            for row in sem_data[sem_name]:
                if len(row) >= 2:
                    key = normalize_key(row[0], row[1])
                    if key not in sem_membership:
                        sem_membership[key] = []
                    if sem_name not in sem_membership[key]:
                        sem_membership[key].append(sem_name)

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
                'apellidos': apellidos,
                'nombre': nombre,
                'email': email,
                'institucion': institucion,
                'pais': pais,
                'seminarios': seminarios,
                'total_seminarios': len(seminarios),
                'key_id': email if email else f"{apellidos.lower()}_{nombre.lower()}",
                'tipo_inscripcion': 'Solo Congreso' if len(seminarios) == 0 else f'Congreso + {len(seminarios)} Seminario(s)'
            })

        attendance = {}
        try:
            ws_att = sh.worksheet('Asistencia')
            att_rows = ws_att.get_all_values()
            if len(att_rows) > 1:
                headers = att_rows[0]
                for arr in att_rows[1:]:
                    if len(arr) >= 2 and arr[0]:
                        kid = arr[0]
                        record = {'seminarios': {}}
                        if len(arr) > 1 and arr[1]:
                            record['timestamp'] = arr[1]
                        if len(arr) > 2:
                            record['recepcionista'] = arr[2]
                        if len(arr) > 3 and arr[3]:
                            record['seminarios']['Seminario 1'] = arr[3]
                        if len(arr) > 4 and arr[4]:
                            record['seminarios']['Seminario 2'] = arr[4]
                        if len(arr) > 5 and arr[5]:
                            record['seminarios']['Seminario 3'] = arr[5]
                        if len(arr) > 6 and arr[6]:
                            record['seminarios']['Seminario 4'] = arr[6]
                        attendance[kid] = record
        except:
            ws_att = sh.add_worksheet('Asistencia', 1000, 10)
            ws_att.append_row(['key_id', 'timestamp', 'recepcionista', 'Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4'])

        return attendees, attendance
    except Exception as e:
        print(f'Error cargando datos: {e}')
        traceback.print_exc()
        return [], {}

def save_attendance_sheet(attendance):
    gc = get_gc()
    if not gc:
        return
    try:
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet('Asistencia')
        ws.clear()
        headers = ['key_id', 'timestamp', 'recepcionista', 'Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']
        ws.append_row(headers)
        rows = []
        for kid, rec in attendance.items():
            sem = rec.get('seminarios', {})
            rows.append([
                kid,
                rec.get('timestamp', ''),
                rec.get('recepcionista', ''),
                sem.get('Seminario 1', ''),
                sem.get('Seminario 2', ''),
                sem.get('Seminario 3', ''),
                sem.get('Seminario 4', '')
            ])
        if rows:
            ws.append_rows(rows)
    except Exception as e:
        print(f'Error guardando asistencia: {e}')

def write_new_attendee_to_sheets(data):
    gc = get_gc()
    if not gc:
        return False
    try:
        sh = gc.open_by_key(SHEET_ID)
        ws_ai = sh.worksheet('Asistentes_imprimir')
        next_num = len(ws_ai.get_all_values())
        ws_ai.append_row([next_num, data['apellidos'], data['nombre'], data.get('institucion', ''), data.get('pais', '')])
        for sem in data.get('seminarios', []):
            try:
                ws_sem = sh.worksheet(sem)
                ws_sem.append_row([data['apellidos'], data['nombre'], data.get('email', ''), data.get('institucion', ''), data.get('pais', '')])
            except:
                pass
        return True
    except Exception as e:
        print(f'Error escribiendo nuevo asistente: {e}')
        return False

attendees, attendance_data = load_data()

@app.route('/')
def index():
    return render_template('index.html', total=len(attendees), asistieron=sum(1 for k in attendance_data if 'timestamp' in attendance_data[k]))

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

@app.route('/api/mark', methods=['POST'])
def mark_attendance():
    global attendance_data
    data = request.json
    key_id = data.get('key_id', '')
    recepcionista = data.get('recepcionista', '')
    if not key_id:
        return jsonify({'error': 'ID requerido'}), 400
    if key_id in attendance_data and 'timestamp' in attendance_data[key_id]:
        return jsonify({'status': 'already'})
    if key_id not in attendance_data:
        attendance_data[key_id] = {'seminarios': {}}
    attendance_data[key_id]['timestamp'] = datetime.now().isoformat()
    attendance_data[key_id]['recepcionista'] = recepcionista
    if 'seminarios' not in attendance_data[key_id]:
        attendance_data[key_id]['seminarios'] = {}
    save_attendance_sheet(attendance_data)
    return jsonify({'status': 'ok', 'message': 'Asistencia registrada'})

@app.route('/api/mark-seminar', methods=['POST'])
def mark_seminar():
    global attendance_data
    data = request.json
    key_id = data.get('key_id', '')
    seminario = data.get('seminario', '')
    if not key_id or not seminario:
        return jsonify({'error': 'ID y seminario requeridos'}), 400
    if key_id not in attendance_data:
        attendance_data[key_id] = {'seminarios': {}}
    if 'seminarios' not in attendance_data[key_id]:
        attendance_data[key_id]['seminarios'] = {}
    if seminario in attendance_data[key_id]['seminarios']:
        return jsonify({'status': 'already'})
    attendance_data[key_id]['seminarios'][seminario] = datetime.now().isoformat()
    save_attendance_sheet(attendance_data)
    return jsonify({'status': 'ok', 'message': f'{seminario} registrado'})

@app.route('/api/status/<path:key_id>')
def get_status(key_id):
    rec = attendance_data.get(key_id, {})
    return jsonify({
        'asistio': 'timestamp' in rec,
        'timestamp': rec.get('timestamp', ''),
        'seminarios': rec.get('seminarios', {})
    })

@app.route('/api/register-last-minute', methods=['POST'])
def register_last_minute():
    global attendees, attendance_data
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
    ok = write_new_attendee_to_sheets(new_a)
    attendees.append(new_a)
    return jsonify({'status': 'ok', 'gsheet_ok': ok, 'total': len(attendees)})

@app.route('/api/stats')
def stats():
    sem_stats = {}
    for sem in ['Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']:
        inscritos = sum(1 for a in attendees if sem in a['seminarios'])
        asistieron_sem = sum(1 for a in attendees if sem in a['seminarios'] and attendance_data.get(a['key_id'], {}).get('seminarios', {}).get(sem))
        sem_stats[sem] = {'inscritos': inscritos, 'asistieron': asistieron_sem}
    marcados = sum(1 for k in attendance_data if 'timestamp' in attendance_data[k])
    solo = sum(1 for a in attendees if a['total_seminarios'] == 0)
    con = sum(1 for a in attendees if a['total_seminarios'] > 0)
    return jsonify({
        'total': len(attendees), 'asistieron': marcados, 'pendientes': len(attendees) - marcados,
        'solo_congreso': solo, 'con_seminarios': con, 'seminarios': sem_stats
    })

@app.route('/api/reload', methods=['POST'])
def reload_data():
    global attendees, attendance_data
    attendees, attendance_data = load_data()
    return jsonify({'status': 'ok', 'total': len(attendees)})

@app.route('/report')
def report():
    return render_template('report.html')

def build_report_data():
    rows = []
    for a in attendees:
        rec = attendance_data.get(a['key_id'], {})
        sem_asistidos = rec.get('seminarios', {}) if rec else {}
        rows.append({
            'nombre_completo': f"{a['nombre']} {a['apellidos']}",
            'apellidos': a['apellidos'], 'nombre': a['nombre'], 'email': a['email'],
            'institucion': a['institucion'], 'pais': a['pais'], 'tipo': a['tipo_inscripcion'],
            'asistio_general': 'timestamp' in rec,
            'seminarios_inscritos': a['seminarios'], 'total_inscritos': a['total_seminarios'],
            'seminarios_asistidos': sem_asistidos, 'total_asistidos': len(sem_asistidos),
            'Seminario 1': 'Seminario 1' in sem_asistidos, 'Seminario 2': 'Seminario 2' in sem_asistidos,
            'Seminario 3': 'Seminario 3' in sem_asistidos, 'Seminario 4': 'Seminario 4' in sem_asistidos,
        })
    return rows

@app.route('/api/report-data')
def report_data():
    rows = build_report_data()
    sem_stats = {}
    for sem in ['Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4']:
        inscritos = sum(1 for a in attendees if sem in a['seminarios'])
        asistieron_sem = sum(1 for a in attendees if sem in a['seminarios'] and attendance_data.get(a['key_id'], {}).get('seminarios', {}).get(sem))
        sem_stats[sem] = {'inscritos': inscritos, 'asistieron': asistieron_sem, 'porcentaje': round(asistieron_sem / inscritos * 100, 1) if inscritos > 0 else 0}
    total_gen = sum(1 for r in rows if r['asistio_general'])
    return jsonify({
        'total': len(rows), 'asistieron_general': total_gen, 'pendientes_general': len(rows) - total_gen,
        'pct_general': round(total_gen / len(rows) * 100, 1) if rows else 0,
        'seminarios': sem_stats, 'rows': rows
    })

@app.route('/api/report-csv')
def report_csv():
    rows = build_report_data()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Apellidos', 'Nombre', 'Email', 'Institucion', 'Pais', 'Tipo Inscripcion',
                     'Asistio General', 'Seminarios Inscritos', 'Seminarios Asistidos',
                     'Seminario 1', 'Seminario 2', 'Seminario 3', 'Seminario 4'])
    for r in rows:
        writer.writerow([r['apellidos'], r['nombre'], r['email'], r['institucion'], r['pais'], r['tipo'],
            'Si' if r['asistio_general'] else 'No', ', '.join(r['seminarios_inscritos']),
            ', '.join(r['seminarios_asistidos'].keys()),
            'Si' if r['Seminario 1'] else 'No', 'Si' if r['Seminario 2'] else 'No',
            'Si' if r['Seminario 3'] else 'No', 'Si' if r['Seminario 4'] else 'No'])
    return Response(output.getvalue(), mimetype='text/csv',
                    headers={'Content-Disposition': 'attachment; filename=reporte_asistencia_cruc_2026.csv'})

@app.route('/api/diagnostic')
def diagnostic():
    gc = get_gc()
    return jsonify({'gc_ok': gc is not None, 'total_asistentes': len(attendees), 'total_asistencias': len(attendance_data)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f'Servidor iniciado en http://localhost:{port}')
    print(f'Asistentes cargados: {len(attendees)}')
    print(f'Asistencias registradas: {len(attendance_data)}')
    app.run(debug=False, host='0.0.0.0', port=port)
