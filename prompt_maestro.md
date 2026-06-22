Actúa como un Ingeniero de Software Senior especialista en automatización de procesos, integraciones de Google Workspace y desarrollo de Pipelines de Inteligencia Artificial en Python. 

Tu tarea es escribir un script de Python completo, funcional, robusto y profesional para un "Sistema de Conciliación de Pagos Institucional". Debes seguir de forma estricta las especificaciones de arquitectura, nombres de variables, lógica de control y manejo de errores detallados a continuación. No omitas código, no uses marcadores de posición (placeholders como "..."), ni resumas funciones.

---

### 1. CONFIGURACIÓN E IMPORTACIONES
El script debe comenzar importando exactamente las siguientes librerías:
- Estándar: `os`, `re`, `io`, `sys`, `json`, `time`, `base64`, `requests`, `datetime`, `email.utils`
- Análisis de datos: `pandas as pd`
- Google API: `google.auth.transport.requests.Request`, `google.oauth2.credentials.Credentials`, `google_auth_oauthlib.flow.Flow`, `googleapiclient.discovery.build`, `googleapiclient.http.MediaIoBaseDownload`, `MediaFileUpload`
- Estilos de Excel: `openpyxl`, `openpyxl.styles.Font`

Establece las siguientes variables de entorno globales obligatorias para OAuthlib:
- `os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'`
- `os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'`

---

### 2. VARIABLES GLOBALES Y CONSTANTES
Define exactamente las siguientes constantes globales:
- `MODELO_IA = "gemini-2.5-flash"`
- `ARCHIVO_EXCEL_LOCAL = "reg_anon.xlsx"`
- `ARCHIVO_MAPEO_LOCAL = "mapeo_identidad.xlsx"`
- `ARCHIVO_REGISTROS_LOCAL = "registros.xlsx"`
- `NOMBRE_EXCEL_DRIVE = "reg_anon.xlsx"`
- `NOMBRE_MAPEO_DRIVE = "mapeo_identidad.xlsx"`
- `NOMBRE_REGISTROS_DRIVE = "registros.xlsx"`
- `RUTA_CREDENTIALS = "credentials.json"`
- `RUTA_TOKEN = "token.json"`
- `ETIQUETA_CONTROL = "CONCILIADO"`
- `SCOPES = ['https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/drive']`
- `MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]`
- `PALABRAS_CLAVE_PAGO = ["PAGO", "COMPROBANTE", "TRANSFERENCIA", "RECIBO", "DEPÓSITO", "DEPOSITO"]`

Recupera de forma segura la API Key del entorno mediante:
- `API_KEY = os.environ.get('GEMINI_API_KEY')`

---

### 3. FUNCIONES AUXILIARES

#### 3.1. `normalizar_id_cliente(valor)`
- Recibe un valor de ID. Si es `pd.isna()`, retorna una cadena vacía `""`.
- En caso contrario, lo convierte a string, limpia espacios, pasa a mayúsculas y, si termina con `'.0'` (debido a conversiones automáticas de Excel/Pandas), elimina ese sufijo (`val_str[:-2]`). Retorna el valor normalizado.

#### 3.2. `obtener_credenciales()`
- Inicializa `creds = None`.
- Si `RUTA_TOKEN` existe, carga las credenciales usando `Credentials.from_authorized_user_file`.
- Si las credenciales no son válidas o están expiradas (y hay un refresh token), intenta renovarlas con `creds.refresh(Request())`. Si falla, imprime una advertencia, requiere reautorización y limpia `creds = None`.
- Si `creds` es `None`, verifica si `RUTA_CREDENTIALS` existe; si no, lanza un `FileNotFoundError`.
- Si existe, usa `Flow.from_client_secrets_file` con los `SCOPES` y el URI `http://localhost` para generar la URL de autorización. Imprime por consola instrucciones claras y formateadas pidiendo al usuario abrir la URL y pegar la URL de localhost resultante.
- Captura la URL usando `input()`, recupera el token mediante `flow.fetch_token()` y guarda las credenciales serializadas a JSON en `RUTA_TOKEN`. Retorna `creds`.

#### 3.3. `descargar_archivo_de_drive(servicio_drive, nombre_archivo, ruta_destino)`
- Busca en Drive un archivo no eliminado con el nombre exacto mediante query `name = '{nombre_archivo}' and trashed = false`.
- Si no hay resultados, imprime error y retorna `None`.
- Si se encuentra, toma el primer `id`, descarga su contenido binario en un búfer de bytes (`io.BytesIO`) iterando bloques con `MediaIoBaseDownload`, y guarda el archivo descargado localmente en `ruta_destino`. Retorna el `file_id`.

#### 3.4. `actualizar_archivo_en_drive(servicio_drive, file_id, ruta_local)`
- Sube el archivo local `ruta_local` de vuelta a Google Drive reemplazando los metadatos del `file_id` mediante `servicio_drive.files().update` y usando `MediaFileUpload` con el mimetype de hoja de cálculo de Excel. Retorna un booleano indicando el éxito de la sincronización.

#### 3.5. `obtener_o_crear_etiqueta(servicio_gmail, nombre_etiqueta)`
- Obtiene la lista de etiquetas del usuario. Si la etiqueta ya existe (insensible a mayúsculas), retorna su `id`.
- Si no existe, la crea con `servicio_gmail.users().labels().create()` configurando visibilidad normal y retorna su nuevo `id`.

#### 3.6. `aplicar_etiqueta_procesado(servicio_gmail, message_id, label_id)`
- Utiliza la API `batchModify` para añadir `label_id` al correo y quitarle la etiqueta `'UNREAD'` (no leído). Imprime un log del proceso.

#### 3.7. `extraer_cuerpo_texto(payload)`
- Extrae de forma recursiva todo el contenido del cuerpo del correo electrónico.
- Si la parte del payload es `text/plain`, decodifica usando UTF-8 ignorando errores.
- Si la parte es `text/html`, decodifica, limpia todas las etiquetas HTML usando expresiones regulares (`r'<[^>]+>'`) y reemplaza entidades comunes de escape como `&nbsp;`, `&amp;` y `&gt;`.
- Elimina espacios acumulados usando `re.sub(r'\s+', ' ', texto_sucio)` y retorna el string resultante.

#### 3.8. `buscar_palabras_clave(asunto, cuerpo)`
- Verifica si en el asunto o cuerpo aparece alguna de las cadenas contenidas en `PALABRAS_CLAVE_PAGO` en mayúsculas. Retorna un booleano.

#### 3.9. `detectar_mes_en_texto(asunto, cuerpo)`
- Limpia de puntuación y tabulaciones el texto completo (asunto + cuerpo). 
- Escanea con expresiones regulares usando límites de palabra (`\b` + mes + `\b`) si se menciona algún mes de la lista `MESES`. Si lo encuentra, retorna el nombre del mes.

#### 3.10. `detectar_mes_en_hilo(servicio_gmail, thread_id)`
- Descarga el hilo completo (`threads().get`) y concatena los asuntos y cuerpos de todos los mensajes previos en ese hilo de conversación para escanearlos usando `detectar_mes_en_texto`.

#### 3.11. `obtener_mes_desde_fecha(fecha_str)`
- Extrae el mes en mayúsculas a partir de un formato `'DD/MM/YYYY'` dividiendo la cadena e indexando contra la lista `MESES` (restando 1).

#### 3.12. `obtener_adjuntos_recursivo(parte)`
- Explora de manera recursiva todas las partes de la estructura del correo electrónico, acumulando y retornando una lista con todos los objetos de parte que representen un archivo adjunto con `filename` y `attachmentId`.

---

### 4. PROCESAMIENTO INTELIGENTE DE METADATOS (IA MULTIMODAL)

#### 4.1. `limpiar_y_cargar_json(texto_respuesta)`
- Sanitiza la cadena de texto de la IA eliminando marcas de bloque Markdown como ` ```json` al principio y ` ``` ` al final de forma insensible a mayúsculas usando expresiones regulares multilinea. Convierte y retorna el objeto parsed mediante `json.loads()`.

#### 4.2. `extraer_metadatos_con_gemini(contenido_binario, mime_type)`
- Valida la existencia de `API_KEY`. Envía la imagen codificada en Base64 en un payload JSON por método POST a la URL de la API REST oficial de Gemini v1beta (`models/gemini-2.5-flash:generateContent`).
- Configura el system prompt para indicarle al modelo actuar como extractor estructurado de recibos de pago de México, solicitando la fecha formateada a `'DD/MM/YYYY'` y el monto total como float, restringiendo la respuesta única al esquema JSON exacto: `{"fecha": "DD/MM/YYYY", "monto": 0.0}`.
- Implementa una política de **Exponential Backoff** de 5 intentos usando una lista de retardos: `delays = [2, 4, 8, 16, 32]`.
- Si el código HTTP es `200`, extrae el texto del candidato, aplica una **pausa preventiva obligatoria de 4 segundos** (`time.sleep(4)`) para respetar el límite de 15 RPM del plan gratuito, sanitiza la respuesta y retorna el JSON mapeado.
- Si la respuesta es un error `429` (cuota) o `503` (saturación), imprime un mensaje detallado del reintento, espera los segundos definidos en `delay` y continúa en bucle. Retorna `None` si se agotan los intentos.

---

### 5. MOTOR DE REGLAS DE CASCADA Y ACTUALIZACIÓN EXCEL

#### 5.1. `actualizar_excel(id_cliente, mes_pago, monto_pago)`
Esta es la función transaccional núcleo. Sigue detalladamente estos pasos:

1. **Actualización local de `reg_anon.xlsx`:**
   - Abre el archivo usando `openpyxl.load_workbook`.
   - Lee la primera fila para mapear la posición exacta de las columnas de los encabezados (pasando a mayúsculas y eliminando espacios).
   - Busca de forma lineal en la fila 1 la columna que coincida con el ID del cliente normalizado mediante `normalizar_id_cliente`. Si no existe, aborta.
   - Valida que el mes solicitado pertenezca a la lista `MESES`.
   - **Lógica de Cascada (Desborde):** Obtiene el índice cronológico del mes en `MESES`. Comienza a recorrer desde ese índice hacia diciembre (`len(MESES)`). Si el mes candidato existe en los encabezados, revisa la celda correspondiente al cliente. Si la celda está vacía (`None`), tiene texto vacío o es numéricamente igual a `0.0`, selecciona ese mes como destino y rompe el ciclo. 
   - Si el destino elegido es cronológicamente posterior al mes original solicitado, marca la bandera `es_pago_adelantado = True`.
   - Si todos los meses posteriores están ocupados, acumula de forma forzada en el último mes registrado del año y marca la bandera `es_pago_adelantado = True`.
   - Obtiene el valor anterior de la celda de destino (considerando `0.0` si es nulo), suma el nuevo monto de pago y asigna el total a la celda.
   - **Formateo de openpyxl:** Si la bandera `es_pago_adelantado` es verdadera, aplica un estilo visual sofisticado a la celda de destino: tipo de fuente `Arial`, tamaño `11`, en **Negrita**, *Cursiva* y color de texto **Verde Hexagonal** (`'008000'`). Si no es adelantado, asigna estilo estándar negro.
   - Guarda los cambios en `reg_anon.xlsx` y cierra el libro.

2. **Sincronización local de `registros.xlsx`:**
   - Lee el archivo `mapeo_identidad.xlsx` con Pandas para encontrar el ID real de la institución asociado al `id_cliente_normalizado` (en la columna de mapeo correspondiente).
   - Si se encuentra el ID real y existe el archivo `registros.xlsx` localmente, abre este último usando `openpyxl.load_workbook`.
   - Mapea de forma lineal sus encabezados, busca la fila del ID real del cliente y la columna del mismo mes final de destino resuelto en la fase de cascada previa.
   - Realiza la misma suma del monto e impone el mismo estilo de celda (Negrita, Cursiva, Verde `'008000'`) si el pago fue redireccionado (adelantado).
   - Guarda los cambios en `registros.xlsx` y cierra el libro.
   - Retorna el nombre del mes en el que se consolidó la transacción final o `None` en caso de fallo crítico.

---

### 6. ORQUESTACIÓN Y FLUJO DE EJECUCIÓN PRINCIPAL (`main`)
- Al iniciar, imprime un banner institucional: `=== SISTEMA INTELIGENTE DE CONCILIACIÓN INSTITUCIONAL ===`.
- Valida que la variable global `API_KEY` esté configurada en el sistema.
- **Paso 1 (Autenticación):** Llama a `obtener_credenciales()` para levantar el objeto OAuth2 y construye las APIs de Gmail (`'gmail'`, `'v1'`) y Drive (`'drive'`, `'v3'`) mediante `googleapiclient.discovery.build`.
- **Paso 2 (Descarga de Drive):** Descarga temporalmente los tres Excels de Google Drive a sus contrapartes locales usando `descargar_archivo_de_drive` (Mapeo, Anónimo y Registros). Si alguno falta, interrumpe el script de forma controlada.
- **Paso 3 (Carga de Mapeo):** Carga con Pandas `mapeo_identidad.xlsx`, normaliza los ID y pasa los correos a minúsculas, filtrando una lista única de correos institucionales registrados de tus clientes.
- **Paso 4 (Gmail Scan):** Busca y crea si es necesario la etiqueta `CONCILIADO` en Gmail. Construye un Query inteligente con operadores `OR` para filtrar únicamente correos que tengan archivos adjuntos, que no contengan ya la etiqueta `CONCILIADO` y que procedan de la lista de correos de clientes registrada.
- **Paso 5 (Pipeline de Transacción):** Para cada correo encontrado:
  - Recupera la metadata del remitente, el asunto y el cuerpo.
  - Verifica si el mensaje trata sobre cobros llamando a `buscar_palabras_clave`.
  - Intenta encontrar el mes correspondiente en el asunto o cuerpo. Si no lo halla, analiza recursivamente el hilo del correo.
  - Explora todos los archivos adjuntos. Si encuentra imágenes (.jpg, .jpeg, .png) o archivos .pdf, procede a enviar los bytes de archivo a la API de Gemini mediante `extraer_metadatos_con_gemini`.
  - Si la IA extrae un monto y fecha de recibo válidos, intenta determinar el mes a partir de la fecha del recibo si aún no se había resuelto.
  - **Manejo de Entrada Fallback:** Si tras todos los análisis no se puede identificar el mes, evalúa si la sesión es interactiva usando `sys.stdin.isatty()`. Si es interactiva, solicita el mes al operador vía consola (`input()`). Si se ejecuta en segundo plano (Cron/VPS), asigna de forma automática y desatendida el mes actual en curso de la máquina.
  - Si se determina un mes definitivo de destino, añade la transacción estructurada a una lista acumuladora.
- **Paso 6 (Actualización y Sincronización final):** Si la lista de transacciones tiene elementos, actualiza para cada una los archivos locales mediante `actualizar_excel` y marca el correo electrónico en Gmail como `'CONCILIADO'` y leídos usando `aplicar_etiqueta_procesado`. Al finalizar, sincroniza y sube los archivos modificados localmente de regreso a sus respectivas ubicaciones en Google Drive. Si no hay elementos, imprime un mensaje informando que no hubo nuevos pagos para conciliar en la ejecución.

Escribe el código completo, estructurado, limpio y bien comentado en español.
