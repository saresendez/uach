# uach
Proyecto UACH -FING

Este proyecto consiste en un Agente Inteligente de Conciliación Bancaria desatendido, diseñado para automatizar el flujo contable de cuentas por cobrar. Combina el Procesamiento de Lenguaje Natural (NLP) y la Visión Computacional Multimodal (LLMs) para eliminar la intervención humana en la lectura, validación y registro de transferencias financieras, garantizando la privacidad de los datos mediante un diseño híbrido.

Características Clave

Visión Computacional Multimodal: Utiliza gemini-2.5-flash para realizar un análisis semántico del comprobante (formatos .pdf, .jpeg, .png), extrayendo el monto neto y la fecha sin depender de plantillas rígidas (coordenadas de píxeles) ni OCR sintáctico tradicional.

Arquitectura Híbrida de Privacidad: Para cumplir con normativas de protección de datos personales, el sistema utiliza un Tokenizador de Identidad local (mapeo_identidad.xlsx). Los datos sensibles (nombres, correos reales) nunca se envían a la nube de Google AI Studio, solo se transmite la imagen del recibo de manera anónima.

Lógica de Cascada y Redirección Inteligente (Smart Allocation): Si un cliente liquida un mes por adelantado, el script detecta cronológicamente el siguiente mes disponible en la base de datos y aplica el pago, dándole formato visual estilizado (Negrita, Cursiva, Color Verde) en el libro Excel mediante openpyxl.

Resiliencia ante Límites de API (Rate-Limiting): Implementa un algoritmo matemático de Backoff Exponencial con Amortiguación para manejar errores de cuota HTTP 429 (Too Many Requests) o saturación del servidor 503.

Despliegue Desatendido 24/7: Configurado para ejecutarse de manera autónoma en un VPS Linux (Ubuntu) a través de tareas programadas (cron), con un sistema robusto de auditoría de logs.

                +----------------------------------------+
               |  Gmail: Escaneo inteligente de Inbox   |
               |  (Filtros: Attachments & Keywords)     |
               +----------------------------------------+
                                   |
                                   v
               +----------------------------------------+
               | Tokenizador Local: Mapeo de Identidad  |  <-- Resguarda Datos
               | (Cruce de correo -> ID_Cliente)        |      Sensibles
               +----------------------------------------+
                                   |
                                   v
               +----------------------------------------+
               |  API Gemini: Extracción Semántica     |  <-- Solo procesa
               |  (Monto neto, Fecha del Recibo)        |      el binario anónimo
               +----------------------------------------+
                                   |
                                   v
               +----------------------------------------+
               |  Motor Contable: Cascada de Meses      |
               |  & Aplicación de Formatos en Caliente  |
               +----------------------------------------+
                                   |
               +-------------------+--------------------+
               |                                        |
               v                                        v
+-----------------------------+          +-----------------------------+
| Actualiza: reg_anon.xlsx    |          | Actualiza: registros.xlsx   |
| (ID Anónimo de Control)     |          | (ID Real del Alumno)        |
+-----------------------------+          +-----------------------------+
               |                                        |
               +-------------------+--------------------+
                                   |
                                   v
               +----------------------------------------+
               |   Google Drive: Sincronización Remota  |
               |   y Etiquetado del correo procesado    |
               +----------------------------------------+

-Requisitos e Instalación

-Requisitos Previos

Servidor VPS Linux( sitema basado en Debian recomendado) o entorno local con Python 3.10+.
Cuenta en Google Cloud Console con las APIs de Gmail y Google Drive habilitadas.
Archivo credentials.json generado desde la pantalla de consentimiento de OAuth de Google.
Una API Key válida de Google AI Studio (Gemini API).

Instalación del Entorno

Clona el repositorio en tu servidor:
___________________________________________________________________
git clone https://github.com/tu-usuario/conciliador-inteligente.git
cd conciliador-inteligente
___________________________________________________________________

Crea y activa el entorno virtual de Python:
____________________________________
python3 -m venv venv
source venv/bin/activate
____________________________________

Instala las dependencias requeridas:
____________________________________
pip install -r requirements.txt
____________________________________

Coloca tus archivos de configuración (credentials.json, mapeo_identidad.xlsx, reg_anon.xlsx, registros.xlsx) dentro de la carpeta raíz de tu proyecto.

-Configuración y Uso

Ejecución Manual
Para realizar pruebas y otorgar el primer consentimiento de OAuth (se abrirá un enlace en consola para que pegues la URL de redirección localhost):
____________________________________________
export GEMINI_API_KEY="tu_clave_de_gemini"
python proyecto.py
____________________________________________

Programación del Servicio Autónomo (Cron Job)
Para dejar el robot trabajando de fondo cada 15 minutos en tu VPS, edita tu tabla de cron:
________________
crontab -e
________________

Añade la siguiente regla al final del archivo (asegúrate de especificar tu ruta absoluta y clave real):
____________________________________________
*/15 * * * * cd /home/conciliador && export GEMINI_API_KEY="TU_API_KEY_AQUI" && /home/conciliador/venv/bin/python proyecto.py >> /home/conciliador/conciliacion.log 2>&1
____________________________________________

-Auditoría y Monitoreo

Puedes revisar el desempeño del conciliador y el resultado de las extracciones semánticas en tiempo real leyendo la bitácora de ejecución del VPS:
___________________________________________
tail -f /home/conciliador/conciliacion.log
___________________________________________
