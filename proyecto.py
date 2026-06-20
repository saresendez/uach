import os
import re
import json
import time
import base64
import requests
import email.utils
import pandas as pd
import sys
import html
from datetime import datetime

# Detección dinámica de entorno (Google Colab vs. GitHub Codespaces / Local PC)
try:
    from google.colab import userdata
    IN_COLAB = True
except ImportError:
    IN_COLAB = False

# Librerías oficiales de Google para OAuth2 y Gmail API
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

# Librerías para dar formato avanzado a celdas de Excel (Color, Negrita, Cursiva)
import openpyxl
from openpyxl.styles import Font, PatternFill

# Permitir que oauthlib acepte URLs HTTP locales (http://localhost) durante el flujo de autenticación
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'

MODELO_IA = "gemini-2.5-flash"

# ==========================================
# CONFIGURACIÓN DINÁMICA DE RUTAS Y API KEY
# ==========================================
if IN_COLAB:
    BASE_DIR = "/content/drive/MyDrive/Proyecto"
else:
    posible_directorio = os.path.join(os.getcwd(), "Proyecto")
    if os.path.exists(posible_directorio):
        BASE_DIR = posible_directorio
    else:
        BASE_DIR = os.getcwd()

ARCHIVO_EXCEL = os.path.join(BASE_DIR, "reg_anon.xlsx")
ARCHIVO_MAPEO = os.path.join(BASE_DIR, "mapeo_identidad.xlsx")
ARCHIVO_REGISTROS = os.path.join(BASE_DIR, "registros.xlsx")
RUTA_CREDENTIALS = os.path.join(BASE_DIR, "credentials.json")
RUTA_TOKEN = os.path.join(BASE_DIR, "token.json")

ETIQUETA_CONTROL = "CONCILIADO"
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Listas de control semántico (Orden cronológico estricto para la cascada de meses)
PALABRAS_CLAVE_PAGO = ["PAGO", "COMPROBANTE", "TRANSFERENCIA", "RECIBO", "DEPÓSITO", "DEPOSITO"]
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

API_KEY = None
try:
    if IN_COLAB:
        API_KEY = userdata.get('GEMINI_API_KEY')
    else:
        API_KEY = os.environ.get('GEMINI_API_KEY')

    if not API_KEY:
        raise ValueError("La API Key 'GEMINI_API_KEY' no está configurada.")
    print(f"✅ API Key de Gemini recuperada. Usando: {MODELO_IA}")
except Exception as e:
    print(f"⚠️ Error de configuración de API Key: {e}")

def normalizar_id_cliente(valor):
    """Normaliza IDs de cliente para evitar discrepancias de tipos de datos (ej. 102 vs 102.0)."""
    if pd.isna(valor):
        return ""
    val_str = str(valor).strip().upper()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

def obtener_servicio_gmail():
    """Autentica al usuario mediante OAuth2 manual compatible con Colab y entornos locales."""
    creds = None
    if os.path.exists(RUTA_TOKEN):
        creds = Credentials.from_authorized_user_file(RUTA_TOKEN, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"⚠️ No se pudo refrescar el token anterior ({e}). Iniciando autenticación desde cero...")
                creds = None
        
        if not creds:
            if not os.path.exists(RUTA_CREDENTIALS):
                raise FileNotFoundError(f"❌ No se encontró el archivo credentials.json en: {RUTA_CREDENTIALS}")

            flow = Flow.from_client_secrets_file(
                RUTA_CREDENTIALS,
                scopes=SCOPES,
                redirect_uri='http://localhost'
            )

            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true'
            )

            print("\n================ AUTORIZACIÓN REQUERIDA ================")
            print("1. Abre este enlace en tu navegador para iniciar sesión con tu cuenta:")
            print(f"\n👉 {authorization_url}\n")
            print("2. Selecciona tu cuenta y acepta los permisos de la aplicación.")
            print("3. Al finalizar, copia la URL completa de localhost (ej. http://localhost...)")
            print("========================================================\n")

            url_redirigida = input("Pega aquí la URL completa de localhost que copiaste: ").strip()

            flow.fetch_token(authorization_response=url_redirigida)
            creds = flow.credentials

        with open(RUTA_TOKEN, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def obtener_o_crear_etiqueta(servicio, nombre_etiqueta):
    """Verifica si existe la etiqueta de control y si no, la crea de forma automatizada."""
    try:
        resultados = servicio.users().labels().list(userId='me').execute()
        etiquetas = resultados.get('labels', [])

        for etiqueta in etiquetas:
            if etiqueta['name'].upper() == nombre_etiqueta.upper():
                return etiqueta['id']

        print(f"🏷️ La etiqueta '{nombre_etiqueta}' no existe. Creándola automáticamente...")
        nueva_etiqueta_body = {
            "name": nombre_etiqueta,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show"
        }
        creada = servicio.users().labels().create(userId='me', body=nueva_etiqueta_body).execute()
        print(f"✅ Etiqueta '{nombre_etiqueta}' creada exitosamente.")
        return creada['id']
    except Exception as e:
        print(f"⚠️ Error al gestionar etiquetas de Gmail: {e}")
        return None

def aplicar_etiqueta_procesado(servicio, message_id, label_id):
    """Marca el correo electrónico como CONCILIADO y le quita la etiqueta de 'No leído' (UNREAD)."""
    try:
        servicio.users().messages().batchModify(
            userId='me',
            body={
                'ids': [message_id],
                'addLabelIds': [label_id],
                'removeLabelIds': ['UNREAD']
            }
        ).execute()
        print(f"   🏷️ Correo {message_id} marcado exitosamente como '{ETIQUETA_CONTROL}' y leído.")
    except Exception as e:
        print(f"   ❌ No se pudo aplicar la etiqueta de conciliación al correo {message_id}: {e}")

def extraer_cuerpo_texto(payload):
    """Extrae recursivamente el contenido de texto (plano o HTML) del correo, limpiando etiquetas."""
    def _extraer(parte):
        mime_type = parte.get('mimeType', '')
        body_data = parte.get('body', {}).get('data', '')
        texto = ""

        if mime_type == 'text/plain' and body_data:
            try:
                texto = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
            except Exception:
                pass
        elif mime_type == 'text/html' and body_data:
            try:
                html_raw = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
                texto = re.sub(r'<[^>]+>', ' ', html_raw)
                texto = html.unescape(texto)
            except Exception:
                pass

        partes = parte.get('parts', [])
        texto_acumulado = texto
        for p in partes:
            texto_acumulado += " " + _extraer(p)
        return texto_acumulado

    cuerpo_sucio = _extraer(payload)
    return re.sub(r'\s+', ' ', cuerpo_sucio).strip()

def buscar_palabras_clave(asunto, cuerpo):
    """Verifica si en el asunto o cuerpo aparece alguna palabra clave de pago."""
    texto_completo = f"{asunto} {cuerpo}".upper()
    return any(palabra in texto_completo for palabra in PALABRAS_CLAVE_PAGO)

def detectar_mes_en_texto(asunto, cuerpo):
    """
    Escanea asunto y cuerpo buscando menciones de meses de forma robusta.
    Evita falsos positivos por orden alfabético/de lista y prioriza el Asunto.
    """
    # 1. Intentar buscar en el Asunto (alta prioridad)
    asunto_limpio = re.sub(r'[.,;:()\-_\t/]', ' ', f" {asunto} ".upper())
    encontrados_asunto = []
    for mes in MESES:
        match = re.search(r'\b' + mes + r'\b', asunto_limpio)
        if match:
            encontrados_asunto.append((match.start(), mes))
    
    if encontrados_asunto:
        encontrados_asunto.sort() # Ordenar por posición física en la cadena
        return encontrados_asunto[0][1]

    # 2. Intentar buscar en el Cuerpo
    cuerpo_limpio = re.sub(r'[.,;:()\-_\t/]', ' ', f" {cuerpo} ".upper())
    encontrados_cuerpo = []
    for mes in MESES:
        match = re.search(r'\b' + mes + r'\b', cuerpo_limpio)
        if match:
            encontrados_cuerpo.append((match.start(), mes))

    if encontrados_cuerpo:
        encontrados_cuerpo.sort()
        return encontrados_cuerpo[0][1]

    return None

def detectar_mes_en_hilo(servicio, thread_id):
    """Obtiene todos los mensajes del hilo de conversación para escanear el historial completo."""
    try:
        thread = servicio.users().threads().get(userId='me', id=thread_id).execute()
        mensajes_hilo = thread.get('messages', [])

        texto_acumulado_hilo = ""
        for m in mensajes_hilo:
            p = m.get('payload', {})
            headers = p.get('headers', [])
            asunto = next((h.get('value', '') for h in headers if h.get('name', '').lower() == 'subject'), "")
            cuerpo = extraer_cuerpo_texto(p)
            texto_acumulado_hilo += f" {asunto} {cuerpo}"

        return detectar_mes_en_texto("", texto_acumulado_hilo)
    except Exception as e:
        print(f"   ⚠️ No se pudo escanear el historial del hilo de conversación: {e}")
    return None

def obtener_mes_desde_fecha(fecha_str):
    """Extrae el mes en español a partir de una fecha con formato DD/MM/YYYY."""
    try:
        partes = fecha_str.split('/')
        if len(partes) == 3:
            mes_idx = int(partes[1])
            if 1 <= mes_idx <= 12:
                return MESES[mes_idx - 1]
    except Exception:
        pass
    return None

def obtener_adjuntos_recursivo(parte):
    """Busca recursivamente todas las partes de un mensaje de Gmail para encontrar archivos adjuntos."""
    adjuntos = []
    filename = parte.get('filename')
    body = parte.get('body', {})
    attachment_id = body.get('attachmentId')

    if filename and attachment_id:
        adjuntos.append(parte)

    partes_hijas = parte.get('parts', [])
    for p in partes_hijas:
        adjuntos.extend(obtener_adjuntos_recursivo(p))
    return adjuntos

def extraer_metadatos_con_gemini(contenido_binario, mime_type):
    """Llama a la API de Gemini usando Schema Enforcement para asegurar un JSON válido nativo."""
    if not API_KEY:
        print("   ❌ Error: No se puede llamar a Gemini porque falta la API Key.")
        return None

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{MODELO_IA}:generateContent?key={API_KEY}"
    archivo_base64 = base64.b64encode(contenido_binario).decode('utf-8')
    
    system_prompt = (
        "Actúa como un extractor de metadatos estructurados especializado en recibos de pago. "
        "Analiza el archivo adjunto y extrae la fecha del pago y el monto total."
    )
    
    payload = {
        "contents": [{"parts": [
            {"text": system_prompt},
            {"inlineData": {"mimeType": mime_type, "data": archivo_base64}}
        ]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": {
                "type": "OBJECT",
                "properties": {
                    "fecha": {
                        "type": "STRING",
                        "description": "Fecha del pago extraída del recibo, convertida estrictamente al formato DD/MM/YYYY."
                    },
                    "monto": {
                        "type": "NUMBER",
                        "description": "Monto total pagado como número flotante sin símbolos de moneda."
                    }
                },
                "required": ["fecha", "monto"]
            }
        }
    }

    delays = [2, 4, 8, 16, 32]
    response = None
    success = False

    for i, delay in enumerate(delays):
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload)
            if response.status_code == 200:
                success = True
                break
            elif response.status_code == 429:
                print(f"   ⏳ [Intento {i+1}/{len(delays)}] Cuota de Gemini excedida (429). Esperando {delay}s...")
            elif response.status_code == 503:
                print(f"   ⏳ [Intento {i+1}/{len(delays)}] Servidor saturado (503). Esperando {delay}s...")
            else:
                print(f"   ⚠️ [Intento {i+1}/{len(delays)}] Error de API inesperado ({response.status_code}). Esperando {delay}s...")
        except Exception as e:
            print(f"   ⚠️ Error de conexión: {e}. Esperando {delay}s...")
        time.sleep(delay)

    if not success:
        return None

    try:
        # Al usar responseSchema, la respuesta es garantizada de ser JSON puro directo
        raw_text = response.json()['candidates'][0]['content']['parts'][0]['text'].strip()
        return json.loads(raw_text)
    except Exception as e:
        print(f"   ❌ Error procesando la respuesta de Gemini: {e}")
        return None

def actualizar_excel(id_cliente, id_real, mes_pago, monto_pago):
    """
    Registra el pago en reg_anon.xlsx y de forma sincronizada en registros.xlsx.
    Optimizado para evitar lecturas de mapeo reiteradas e incluye tolerancia a bloqueos de archivos abiertos.
    """
    try:
        # -------------------------------------------------------------
        # PARTE 1: ACTUALIZACIÓN EN reg_anon.xlsx
        # -------------------------------------------------------------
        if not os.path.exists(ARCHIVO_EXCEL):
            print(f"❌ Error: El archivo {ARCHIVO_EXCEL} no existe.")
            return False

        try:
            wb = openpyxl.load_workbook(ARCHIVO_EXCEL)
        except PermissionError:
            print(f"❌ Error de permisos: No se puede abrir '{ARCHIVO_EXCEL}'. Asegúrate de cerrarlo en Excel.")
            return False

        ws = wb.active

        encabezados = [str(ws.cell(row=1, column=col).value).strip().upper() for col in range(1, ws.max_column + 1)]
        id_cliente_normalizado = normalizar_id_cliente(id_cliente)
        
        fila_cliente = None
        for row in range(2, ws.max_row + 1):
            val_celda = normalizar_id_cliente(ws.cell(row=row, column=1).value)
            if val_celda == id_cliente_normalizado:
                fila_cliente = row
                break

        if not fila_cliente:
            print(f"❌ El ID de cliente '{id_cliente_normalizado}' no existe en reg_anon.xlsx.")
            wb.close()
            return False

        mes_pago_upper = str(mes_pago).strip().upper()
        if mes_pago_upper not in MESES:
            print(f"❌ El mes '{mes_pago_upper}' no es un mes cronológico válido.")
            wb.close()
            return False

        idx_mes_inicial = MESES.index(mes_pago_upper)
        mes_destino_final = None
        columna_destino_idx = None
        es_pago_adelantado = False

        for i in range(idx_mes_inicial, len(MESES)):
            mes_candidato = MESES[i]
            if mes_candidato in encabezados:
                col_idx = encabezados.index(mes_candidato) + 1
                valor_celda = ws.cell(row=fila_cliente, column=col_idx).value

                if valor_celda is None or str(valor_celda).strip() == "" or float(valor_celda or 0) == 0.0:
                    mes_destino_final = mes_candidato
                    columna_destino_idx = col_idx
                    if i > idx_mes_inicial:
                        es_pago_adelantado = True
                    break

        if not mes_destino_final:
            print(f"⚠️ Todos los meses posteriores están llenos. Se acumulará en el último mes registrado.")
            for i in reversed(range(len(MESES))):
                mes_candidato = MESES[i]
                if mes_candidato in encabezados:
                    mes_destino_final = mes_candidato
                    columna_destino_idx = encabezados.index(mes_candidato) + 1
                    es_pago_adelantado = True
                    break

        celda_objetivo = ws.cell(row=fila_cliente, column=columna_destino_idx)
        valor_actual = celda_objetivo.value

        try:
            monto_previo = float(valor_actual) if valor_actual is not None else 0.0
        except ValueError:
            monto_previo = 0.0

        nuevo_monto_total = monto_previo + monto_pago
        celda_objetivo.value = nuevo_monto_total

        if es_pago_adelantado:
            print(f"🔄 El mes {mes_pago_upper} ya estaba liquidado. Redireccionando pago a {mes_destino_final}...")
            fuente_especial = Font(name='Arial', size=11, bold=True, italic=True, color='008000')
            celda_objetivo.font = fuente_especial
        else:
            celda_objetivo.font = Font(name='Arial', size=11, bold=False, italic=False, color='000000')

        try:
            wb.save(ARCHIVO_EXCEL)
        except PermissionError:
            print(f"❌ Error de escritura: No se puede guardar '{ARCHIVO_EXCEL}'. Asegúrate de cerrarlo en Excel.")
            wb.close()
            return False
        
        wb.close()
        print(f"✅ Registrado en reg_anon.xlsx para {id_cliente_normalizado} en {mes_destino_final}: ${nuevo_monto_total:.2f}")

        # -------------------------------------------------------------
        # PARTE 2: SINCRONIZACIÓN EN registros.xlsx
        # -------------------------------------------------------------
        if id_real and os.path.exists(ARCHIVO_REGISTROS):
            try:
                try:
                    wb_reg = openpyxl.load_workbook(ARCHIVO_REGISTROS)
                except PermissionError:
                    print(f"❌ Error de permisos: No se puede abrir '{ARCHIVO_REGISTROS}'. Asegúrate de cerrarlo en Excel.")
                    return False

                ws_reg = wb_reg.active
                encabezados_reg = [str(ws_reg.cell(row=1, column=col).value).strip().upper() for col in range(1, ws_reg.max_column + 1)]

                fila_real = None
                for row in range(2, ws_reg.max_row + 1):
                    val_celda = str(ws_reg.cell(row=row, column=1).value).strip().upper()
                    if val_celda.endswith('.0'):
                        val_celda = val_celda[:-2]
                    if val_celda == id_real:
                        fila_real = row
                        break

                if fila_real and mes_destino_final in encabezados_reg:
                    col_reg_idx = encabezados_reg.index(mes_destino_final) + 1
                    celda_reg = ws_reg.cell(row=fila_real, column=col_reg_idx)

                    valor_actual_reg = celda_reg.value
                    try:
                        monto_previo_reg = float(valor_actual_reg) if valor_actual_reg is not None else 0.0
                    except ValueError:
                        monto_previo_reg = 0.0

                    nuevo_monto_reg = monto_previo_reg + monto_pago
                    celda_reg.value = nuevo_monto_reg

                    if es_pago_adelantado:
                        celda_reg.font = Font(name='Arial', size=11, bold=True, italic=True, color='008000')
                    else:
                        celda_reg.font = Font(name='Arial', size=11, bold=False, italic=False, color='000000')

                    try:
                        wb_reg.save(ARCHIVO_REGISTROS)
                        print(f"✅ Sincronizado en registros.xlsx para ID Real '{id_real}' en {mes_destino_final}: ${nuevo_monto_reg:.2f}")
                    except PermissionError:
                        print(f"❌ Error de escritura: No se puede guardar '{ARCHIVO_REGISTROS}'. Asegúrate de cerrarlo.")
                        wb_reg.close()
                        return False
                else:
                    if not fila_real:
                        print(f"⚠️ Advertencia: No se encontró la fila para el ID real '{id_real}' en registros.xlsx.")
                    else:
                        print(f"⚠️ Advertencia: El mes '{mes_destino_final}' no coincide con los encabezados de registros.xlsx.")

                wb_reg.close()
            except Exception as e:
                print(f"❌ Error al intentar escribir en registros.xlsx: {e}")
                return False
        else:
            if not id_real:
                print(f"⚠️ No se encontró ID real en el mapa de identidad para el cliente {id_cliente_normalizado}.")
            if not os.path.exists(ARCHIVO_REGISTROS):
                print(f"⚠️ No se encontró el archivo de registros en la ruta: {ARCHIVO_REGISTROS}")

        return True
    except Exception as e:
        print(f"❌ Error crítico en el guardado transaccional de datos: {e}")
        return False

def procesar_y_conciliar_correos():
    """
    Escanea Gmail, procesa cada correo de forma transaccional, actualiza Excel
    y etiqueta el correo como CONCILIADO de inmediato si todas las operaciones del mismo tienen éxito.
    """
    if not os.path.exists(ARCHIVO_MAPEO):
        print(f"❌ Error crítico: No se encontró el archivo de mapeo en {ARCHIVO_MAPEO}")
        return

    try:
        print("📖 Cargando archivo de mapeo de identidad...")
        df_mapeo = pd.read_excel(ARCHIVO_MAPEO)
        for col in df_mapeo.columns:
            df_mapeo[col] = df_mapeo[col].astype(str).str.strip()

        columna_email = None
        for col in df_mapeo.columns:
            if df_mapeo[col].astype(str).str.contains('@', na=False).any():
                columna_email = col
                break

        if columna_email is None:
            columna_email = df_mapeo.columns[3] if len(df_mapeo.columns) > 3 else df_mapeo.columns[-1]

        columna_id = df_mapeo.columns[0]
        df_mapeo[columna_id] = df_mapeo[columna_id].apply(normalizar_id_cliente)
        df_mapeo[columna_email] = df_mapeo[columna_email].str.lower()
        print(f"   ℹ️ Columnas de mapeo cargadas: ID -> '{columna_id}', Email -> '{columna_email}'")
    except Exception as e:
        print(f"❌ Error al leer mapeo_identidad.xlsx: {e}")
        return

    try:
        servicio = obtener_servicio_gmail()
        label_id = obtener_o_crear_etiqueta(servicio, ETIQUETA_CONTROL)

        query = f'has:attachment -label:{ETIQUETA_CONTROL}'
        print(f"🔍 Escaneando Gmail para nuevos comprobantes: {query}")

        resultado_busqueda = servicio.users().messages().list(userId='me', q=query).execute()
        mensajes = resultado_busqueda.get('messages', [])

        if not mensajes:
            print("📩 No se encontraron nuevos correos con archivos adjuntos pendientes.")
            return

        print(f"📩 Se detectaron {len(mensajes)} correos con adjuntos para inspeccionar.")

        for msg_info in mensajes:
            msg_id = msg_info['id']
            msg = servicio.users().messages().get(userId='me', id=msg_id).execute()
            thread_id = msg.get('threadId')
            payload = msg.get('payload', {})

            headers = payload.get('headers', [])
            from_header = next((h.get('value', '') for h in headers if h.get('name', '').lower() == 'from'), "")
            realname, email_remitente = email.utils.parseaddr(from_header)
            email_remitente = email_remitente.strip().lower()

            match_cliente = df_mapeo[df_mapeo[columna_email] == email_remitente]
            if match_cliente.empty:
                print(f"⚠️ Correo {msg_id} de {email_remitente} omitido: No registrado en el mapa de identidad.")
                continue

            client_id = match_cliente.iloc[0][columna_id]
            
            # Obtener ID real directamente de la segunda columna del mapeo y evitar leer el Excel otra vez
            id_real = str(match_cliente.iloc[0, 1]).strip().upper()
            if id_real.endswith('.0'):
                id_real = id_real[:-2]

            asunto_correo = next((h.get('value', '') for h in headers if h.get('name', '').lower() == 'subject'), "Sin Asunto")
            cuerpo_correo = extraer_cuerpo_texto(payload)

            if not buscar_palabras_clave(asunto_correo, cuerpo_correo):
                continue

            print(f"\n────────────────────────────────────────────────────────")
            print(f"👤 Remitente: {email_remitente} (Cliente ID: {client_id} | Real ID: {id_real})")
            print(f"📧 Asunto: {asunto_correo}")

            mes_pago = detectar_mes_en_texto(asunto_correo, cuerpo_correo)

            if not mes_pago and thread_id:
                print(f"🔍 Mes no detectado en correo actual. Escaneando hilo de conversación...")
                mes_pago = detectar_mes_en_hilo(servicio, thread_id)
                if mes_pago:
                    print(f"💬 Mes hallado en el historial de conversación: {mes_pago}")

            adjuntos = obtener_adjuntos_recursivo(payload)
            
            # Procesar transaccionalmente este correo electrónico
            todo_exitoso = True
            adjuntos_validos = 0
            
            for part in adjuntos:
                nombre_archivo = part.get('filename')
                if nombre_archivo:
                    ext = os.path.splitext(nombre_archivo)[1].lower()
                    mime_type = None
                    if ext in ['.jpg', '.jpeg']: mime_type = "image/jpeg"
                    elif ext == '.png': mime_type = "image/png"
                    elif ext == '.pdf': mime_type = "application/pdf"

                    if mime_type:
                        adjuntos_validos += 1
                        print(f"   📎 Procesando adjunto: {nombre_archivo}")
                        body = part.get('body', {})
                        attachment_id = body.get('attachmentId')

                        if attachment_id:
                            adjunto_raw = servicio.users().messages().attachments().get(
                                userId='me', messageId=msg_id, id=attachment_id).execute()

                            datos_bytes = base64.urlsafe_b64decode(adjunto_raw['data'].encode('UTF-8'))

                            resultado = extraer_metadatos_con_gemini(datos_bytes, mime_type)
                            if resultado and resultado.get('monto') is not None:
                                monto = float(resultado['monto'])
                                
                                # Evitar registrar valores nulos o negativos que no sean pagos reales
                                if monto <= 0:
                                    print(f"      ⚠️ Monto detectado menor o igual a cero (${monto:.2f}). Se omite este adjunto.")
                                    continue
                                
                                fecha_recibo = resultado.get('fecha', '')
                                print(f"      ✨ Extraído -> Fecha Recibo: {fecha_recibo} | Monto: ${monto:.2f}")

                                mes_final = mes_pago
                                if not mes_final:
                                    mes_final = obtener_mes_desde_fecha(fecha_recibo)
                                    if mes_final:
                                        print(f"      📄 Mes determinado por fecha del recibo: {mes_final}")

                                if not mes_final:
                                    if sys.stdin.isatty():
                                        print(f"      ❓ No se pudo identificar el mes automáticamente.")
                                        while True:
                                            entrada = input(f"      👉 Introduce el mes para este pago (ej. MARZO) o ENTER para omitir: ").strip().upper()
                                            if not entrada:
                                                print("      ⏭️ Correo omitido por el usuario.")
                                                break
                                            if entrada in MESES:
                                                mes_final = entrada
                                                break
                                            else:
                                                print(f"      ⚠️ '{entrada}' no es un mes válido. Inténtalo de nuevo.")
                                    else:
                                        # Fallback automático y no bloqueante si corre en servidor sin terminal interactiva
                                        mes_actual_idx = datetime.now().month
                                        mes_final = MESES[mes_actual_idx - 1]
                                        print(f"      📅 Entorno no interactivo. Usando mes actual del sistema: {mes_final}")

                                if mes_final:
                                    # Intentar registrar en Excel de inmediato
                                    escritura_ok = actualizar_excel(client_id, id_real, mes_final, monto)
                                    if not escritura_ok:
                                        todo_exitoso = False
                                        print(f"      ❌ Error al actualizar bases de datos para el adjunto: {nombre_archivo}")
                                else:
                                    todo_exitoso = False
                            else:
                                print(f"      ⚠️ No se pudo extraer información estructurada del adjunto.")
                                todo_exitoso = False

            # Solo marcar el correo como conciliado si hubo adjuntos válidos procesados y no hubo ningún fallo
            if todo_exitoso and adjuntos_validos > 0:
                aplicar_etiqueta_procesado(servicio, msg_id, label_id)
            elif adjuntos_validos == 0:
                print("   ⚠️ El correo no contenía adjuntos compatibles (JPG, PNG o PDF).")
            else:
                print(f"   ⚠️ Se omitió el etiquetado del correo {msg_id} debido a un error durante el flujo de registro.")

    except Exception as e:
        print(f"❌ Error crítico en la API de Gmail: {e}")

if __name__ == "__main__":
    print("=== SISTEMA INTELIGENTE DE CONCILIACIÓN INSTITUCIONAL ===")

    if IN_COLAB:
        print("💻 Ejecutando en Google Colab.")
        if not os.path.exists('/content/drive'):
            from google.colab import drive
            drive.mount('/content/drive')
    else:
        print(f"💻 Ejecutando en entorno local. Carpeta raíz: {BASE_DIR}")

    procesar_y_conciliar_correos()
