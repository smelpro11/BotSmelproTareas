import requests
import pandas as pd
import datetime
import time
import pytz
import json
import os

# Token del bot
TOKEN = "8227348236:AAE5e-s90zqlBujgfpLLZd4h4UFfqY1p1NU"
URL =f"https://api.telegram.org/bot{TOKEN}/sendMessage"
CHAT_ID = -1003343414449
TZ = pytz.timezone('America/Lima')  # Ajusta tu zona horaria

# Archivo para persistir mensajes enviados
ENVIADOS_FILE = "tareas_enviadas.json"

def cargar_enviados():
    """Carga el registro de tareas ya enviadas"""
    if os.path.exists(ENVIADOS_FILE):
        with open(ENVIADOS_FILE, 'r') as f:
            return json.load(f)
    return []

def guardar_enviado(tarea_id):
    """Guarda que una tarea fue enviada"""
    enviados = cargar_enviados()
    if tarea_id not in enviados:
        enviados.append(tarea_id)
        with open(ENVIADOS_FILE, 'w') as f:
            json.dump(enviados, f)

def send_message(text):
    try:
        r = requests.post(URL, json={"chat_id": CHAT_ID, "text": text})
        print(f"[{datetime.datetime.now(TZ)}] Mensaje enviado -> {text}")
        print("Respuesta Telegram:", r.text)
        return True
    except Exception as e:
        print(f"Error enviando mensaje: {e}")
        return False

print("Cargando tareas desde Excel...")
df = pd.read_excel("Tareas.xlsx")

tareas_programadas = []
enviados_previos = cargar_enviados()

for idx, row in df.iterrows():
    empleado = str(row['ID']).strip()
    tarea = str(row['TAREA']).strip()
    
    # Procesar FECHA
    fecha_val = row["FECHA"]
    if isinstance(fecha_val, datetime.datetime):
        fecha = fecha_val.date()
    elif isinstance(fecha_val, datetime.date):
        fecha = fecha_val
    else:
        fecha = datetime.datetime.strptime(str(fecha_val), "%d/%m/%Y").date()
    
    # Procesar HORA
    hora_val = row["HORA"]
    if isinstance(hora_val, datetime.datetime):
        hora = hora_val.time()
    elif isinstance(hora_val, datetime.time):
        hora = hora_val
    elif isinstance(hora_val, pd.Timestamp):
        hora = hora_val.time()
    else:
        hora = datetime.datetime.strptime(str(hora_val), "%H:%M").time()
    
    # Combinar fecha + hora
    fecha_dt = datetime.datetime.combine(fecha, hora)
    fecha_dt = TZ.localize(fecha_dt)
    
    # Crear ID único para la tarea
    tarea_id = f"{empleado}_{tarea}_{fecha_dt.isoformat()}"
    
    mensaje = f"Hola {empleado}, tu tarea para hoy: {tarea} 🧹"
    
    tareas_programadas.append({
        "id": tarea_id,
        "datetime": fecha_dt,
        "mensaje": mensaje,
        "enviado": tarea_id in enviados_previos
    })

# Ordenar por fecha
tareas_programadas.sort(key=lambda x: x['datetime'])

print(f"\nBot activo. {len(tareas_programadas)} tareas cargadas.")
print(f"Tareas ya enviadas previamente: {len(enviados_previos)}")
print(f"Hora actual: {datetime.datetime.now(TZ)}\n")

# Mostrar próximas tareas pendientes
print("=== PRÓXIMAS TAREAS PENDIENTES ===")
now = datetime.datetime.now(TZ)
for tarea in tareas_programadas[:10]:  # Mostrar solo las primeras 10
    if not tarea["enviado"] and tarea["datetime"] > now:
        tiempo_restante = tarea["datetime"] - now
        print(f"  • {tarea['datetime'].strftime('%d/%m %H:%M')} - {tarea['mensaje'][:50]}... (en {tiempo_restante})")
print("===================================\n")

# Ventana de tolerancia: solo enviar si falta menos de este tiempo
VENTANA_MINUTOS = 5

while True:
    now = datetime.datetime.now(TZ)
    
    for tarea in tareas_programadas:
        if not tarea["enviado"]:
            tiempo_diff = (now - tarea["datetime"]).total_seconds() / 60  # diferencia en minutos
            
            # Solo enviar si:
            # 1. La hora programada ya pasó (tiempo_diff >= 0)
            # 2. Pero no hace más de VENTANA_MINUTOS (evita enviar tareas muy antiguas)
            if 0 <= tiempo_diff <= VENTANA_MINUTOS:
                print(f"\n>>> EJECUTANDO TAREA <<<")
                print(f"Programada: {tarea['datetime']}")
                print(f"Ejecutando: {now}")
                print(f"Diferencia: {tiempo_diff:.1f} minutos")
                
                if send_message(tarea["mensaje"]):
                    tarea["enviado"] = True
                    guardar_enviado(tarea["id"])
                    print("✓ Tarea marcada como enviada y guardada\n")
    
    time.sleep(30)  # Revisar cada 30 segundos