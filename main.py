import requests
import pandas as pd
import datetime
import time
import pytz
import json
import os
import sys
import logging

# Configurar logging para que aparezca en journalctl
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Token del bot
TOKEN = "8227348236:AAE5e-s90zqlBujgfpLLZd4h4UFfqY1p1NU"
URL =f"https://api.telegram.org/bot{TOKEN}/sendMessage"
CHAT_ID = -1003343414449
# CHAT_ID = 6197999828
TZ = pytz.timezone('America/Lima')  # Ajusta la zona horaria

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
        logger.info(f"Mensaje enviado -> {text}")
        logger.info(f"Respuesta Telegram: {r.text}")
        return True
    except Exception as e:
        logger.error(f"Error enviando mensaje: {e}")
        return False
    
def verificar_y_enviar_mensajes_recurrentes():
    now = datetime.datetime.now(TZ)
    logger.info(f"Verificando mensajes recurrentes - Día: {now.strftime('%A')} Hora: {now.strftime('%H:%M')}")
    if now.weekday() > 4:  # Solo lunes a viernes
        logger.info("Es fin de semana, no se envían mensajes recurrentes")
        return
    
    fecha_key = now.strftime("%Y-%m-%d")
    horas_envio = [
        {"hora": datetime.time(9, 0), "clave": "09:00", "mensaje": " ¡Hola, equipo! No olviden revisar sus pendientes, priorizar las tareas más importantes y consultar cualquier duda que tengan. Mantengamos el enfoque y evitemos distracciones. ¡Que sea un día lleno de energía, aprendizaje y trabajos muy bien realizados! 💪✨"},
        {"hora": datetime.time(18, 0), "clave": "18:00", "mensaje": "Hoy estamos teniendo un excelente día, lleno de nuevos retos y oportunidades. Antes de retirarte, por favor actualiza tus actividades. Esto te ayudará a organizar mejor tus tareas y a ejecutarlas de manera más eficiente. No olvides limpiar y organizar tu espacio de trabajo antes de salir, ya que un entorno ordenado contribuye a un mejor desempeño. ¡Sigamos avanzando juntos!"}
    ]
    
    enviados = cargar_enviados()
    
    for config in horas_envio:
        mensaje_key = f"recurrente_{fecha_key}_{config['clave']}"
        if mensaje_key in enviados:
            continue
        
        diferencia_minutos = (now.hour * 60 + now.minute) - (config['hora'].hour * 60 + config['hora'].minute)
        if 0 <= diferencia_minutos <= 5:
            logger.info(f">>> ENVIANDO MENSAJE RECURRENTE {config['clave']} <<<")
            if send_message(config['mensaje']):
                guardar_enviado(mensaje_key)

# Forzar flush de stdout
sys.stdout.flush()

logger.info("=== BOT DE TAREAS INICIANDO ===")
logger.info("Cargando tareas desde Excel...")

try:
    df = pd.read_excel("Tareas.xlsx")
    logger.info(f"Excel cargado correctamente. {len(df)} filas encontradas.")
except Exception as e:
    logger.error(f"ERROR al cargar Excel: {e}")
    sys.exit(1)

tareas_programadas = []
enviados_previos = cargar_enviados()

for idx, row in df.iterrows():
    try:
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
    except Exception as e:
        logger.error(f"Error procesando fila {idx}: {e}")
        continue

# Ordenar por fecha
tareas_programadas.sort(key=lambda x: x['datetime'])

logger.info(f"Bot activo. {len(tareas_programadas)} tareas cargadas.")
logger.info(f"Tareas ya enviadas previamente: {len(enviados_previos)}")
logger.info(f"Hora actual: {datetime.datetime.now(TZ)}")

# Mostrar próximas tareas pendientes
logger.info("=== PRÓXIMAS TAREAS PENDIENTES ===")
now = datetime.datetime.now(TZ)

count = 0
for tarea in tareas_programadas:
    if not tarea["enviado"] and tarea["datetime"] > now:
        tiempo_restante = tarea["datetime"] - now
        logger.info(f"  • {tarea['datetime'].strftime('%d/%m %H:%M')} - {tarea['mensaje'][:50]}... (en {tiempo_restante})")
        count += 1
        if count >= 10:
            break

if count == 0:
    logger.warning("¡No hay tareas pendientes futuras!")

logger.info("===================================")

# Ventana de tolerancia: solo enviar si falta menos de este tiempo
VENTANA_MINUTOS = 5

logger.info("Entrando en loop principal...")
sys.stdout.flush()

while True:
    now = datetime.datetime.now(TZ)
    verificar_y_enviar_mensajes_recurrentes()
    for tarea in tareas_programadas:
        if not tarea["enviado"]:
            tiempo_diff = (now - tarea["datetime"]).total_seconds() / 60
            
            if 0 <= tiempo_diff <= VENTANA_MINUTOS:
                logger.info(">>> EJECUTANDO TAREA <<<")
                logger.info(f"Programada: {tarea['datetime']}")
                logger.info(f"Ejecutando: {now}")
                logger.info(f"Diferencia: {tiempo_diff:.1f} minutos")
                
                if send_message(tarea["mensaje"]):
                    tarea["enviado"] = True
                    guardar_enviado(tarea["id"])
                    logger.info("Tarea marcada como enviada y guardada")
    
    time.sleep(30)