# Rendiciones de gastos — app web adaptable a celular

Aplicación web personal para fotografiar varios comprobantes, extraer los datos con IA, revisarlos y generar el PDF con la plantilla Excel oficial.

## Funciones
- Interfaz adaptable a celular y computador.
- Puede instalarse como aplicación web (PWA) en dispositivos compatibles cuando se publique bajo HTTPS.
- Selección múltiple de fotografías o uso de la cámara del teléfono.
- Máximo **39 comprobantes por rendición**.
- Extracción IA exclusiva de: fecha, número de documento, descripción y total final.
- Edición manual de esos cuatro campos.
- La columna Contabilidad de la plantilla siempre queda vacía.
- Aviso de posibles números de documento duplicados.
- Resumen de gastos en vivo y vista previa del PDF antes de descargar.
- PDF configurado en tamaño carta, orientación vertical y ajuste a una sola página.
- La plantilla Excel original se utiliza como base; se conservan encabezados, distribución y espacios de firma.

## Plantilla
`templates/Rendicion Gastos.xlsx`

La hoja `r. d .gastos` se completa así:
- D3: nombre de quien rinde
- F3: fecha de rendición
- B5:B43: fecha del comprobante (39 líneas)
- C5:C43: número de documento
- D5:D43: descripción
- E5:E43: total final
- F5:F43: siempre vacío (Contabilidad)
- E47: suma de los 39 posibles gastos

El PDF se ajusta a una sola hoja carta. Las descripciones largas usan ajuste de texto y reducción automática para evitar que se salgan de sus celdas. Con 39 líneas el contenido puede reducirse de escala para caber en una página; se recomienda revisar la vista previa antes de imprimir.

## Requisitos del servidor
- Python 3.10 o superior
- LibreOffice instalado
- Clave de OpenAI API para reconocimiento de imágenes

## Instalación en Windows
1. Instala Python desde https://www.python.org/downloads/.
2. Instala LibreOffice desde https://www.libreoffice.org/.
3. Descomprime este ZIP.
4. Abre PowerShell dentro de la carpeta del proyecto y ejecuta:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:OPENAI_API_KEY="TU_CLAVE_API"
$env:OPENAI_MODEL="gpt-4.1-mini"
python app.py
```

5. En el computador abre `http://localhost:5000`.

## Uso desde el teléfono en la misma red
- Conecta el teléfono y el computador/servidor a la misma red Wi-Fi.
- Averigua la IP local del computador (por ejemplo, `192.168.1.25`).
- Desde el teléfono abre `http://IP-DEL-COMPUTADOR:5000`.
- Para instalarla como PWA en la pantalla de inicio, publica la aplicación en un dominio con HTTPS. Los navegadores normalmente no permiten instalar PWA desde una IP local HTTP.

## Seguridad y operación
- La clave API se configura como variable de entorno en el servidor; nunca se incluye en el navegador ni en el ZIP.
- Las imágenes se envían al servidor para su análisis y no se guardan en una biblioteca de comprobantes por esta versión.
- La app no está publicada en Internet ni desplegada en un servidor administrado: este ZIP contiene el proyecto para ejecutar o desplegar.
- Antes de imprimir o contabilizar, revisa los cuatro campos. La IA puede equivocarse con imágenes borrosas, inclinadas o cortadas.
- La vista previa PDF requiere LibreOffice en el servidor.
