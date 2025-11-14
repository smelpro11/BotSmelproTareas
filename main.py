import datetime
import time
import requests
import pytz
import pandas as pd

# Token del bot
TOKEN = "8227348236:AAE5e-s90zqlBujgfpLLZd4h4UFfqY1p1NU"
URL = f"https://api.telegram.org/bot{TOKEN}/sendMessage"

# Zona horaria Perú
TZ = pytz.timezone("America/Lima")

# Chat ID único
CHAT_ID = -665612637

def send_message(text):
    try:
        r = requests.post(URL, json={"chat_id": CHAT_ID, "text": text})
        print("Mensaje enviado ->", text)
        print("Respuesta Telegram:", r.text)
    except Exception as e:
        print("Error enviando mensaje:", e)

print("Cargando tareas desde Excel...")

df = pd.read_excel("Tareas.xlsx")
print("\n===== CONTENIDO EXCEL (DEBUG) =====")
print(df)
print("===================================\n")

tareas_programadas = []

for _, row in df.iterrows():
    empleado = str(row['ID']).strip()
    tarea = str(row['TAREA']).strip()

    # ------- FECHA -------
    fecha_val = row["FECHA"]
    print("DEBUG FECHA RAW:", fecha_val, type(fecha_val))

    if isinstance(fecha_val, datetime.datetime):
        fecha = fecha_val.date()
    elif isinstance(fecha_val, datetime.date):
        fecha = fecha_val
    else:
        fecha = datetime.datetime.strptime(str(fecha_val), "%d/%m/%Y").date()

    # ------- HORA -------
    hora_val = row["HORA"]
    print("DEBUG HORA RAW:", hora_val, type(hora_val))

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

    print(f"--> Tarea generada: {empleado} | {tarea} | Fecha y hora final: {fecha_dt}\n")

    mensaje = f"Hola {empleado}, tu tarea para hoy: {tarea} 🧹"

    tareas_programadas.append({
        "datetime": fecha_dt,
        "mensaje": mensaje,
        "enviado": False
    })

print("Bot activo. Esperando tareas...\n")

while True:
    now = datetime.datetime.now(TZ)
    print("Hora actual:", now)

    for tarea in tareas_programadas:
        print(f" Revisando -> {tarea['mensaje']} | programado para {tarea['datetime']} | enviado: {tarea['enviado']}")

        if not tarea["enviado"] and now >= tarea["datetime"]:
            print("\n>>> Ejecutando tarea ahora <<<")
            send_message(tarea["mensaje"])
            tarea["enviado"] = True

    print("Esperando 30 segundos...\n")
    time.sleep(30)
