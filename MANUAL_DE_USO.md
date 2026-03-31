# Manual de Usuario - Bot de Respaldo Quipux

Este bot automatiza la descarga de documentos y anexos desde el sistema Quipux de la ESPE, organizándolos por carpetas y generando un reporte en Excel.

## 🛠️ Requisitos Previos

Asegúrate de tener instalado lo siguiente en tu computadora:

1.  **Python 3.8 o superior**: Puedes descargarlo desde [python.org](https://www.python.org/).
2.  **Bibliotecas de Python**: Instaladas mediante `pip`.
3.  **Navegador Chromium**: Instalado a través de Playwright.

## 🚀 Instalación

Sigue estos pasos para configurar el entorno:

1.  **Instalar las dependencias de Python**:
    Abre una terminal y ejecuta:
    ```bash
    pip install playwright openpyxl
    ```

2.  **Instalar el navegador necesario**:
    Ejecuta el siguiente comando para descargar Chromium:
    ```bash
    playwright install chromium
    ```

## 📋 Uso de la Aplicación

1.  **Ejecutar el Bot**:
    En la terminal, navega hasta la carpeta del proyecto y ejecuta:
    ```bash
    python bot_quipux.py
    ```

2.  **Inicio de Sesión (Manual)**:
    - Se abrirá una ventana del navegador de forma automática.
    - Ingresa tu usuario y contraseña de Quipux.
    - Una vez que veas la bandeja de entrada (Recibidos), regresa a la terminal y presiona **ENTER**.

3.  **Menú de Opciones**:
    El bot presentará un menú interactivo en la terminal:
    - `[1] 📥 Descargar Recibidos`: Procesa y descarga todos los documentos de la bandeja de recibidos.
    - `[2] 📤 Descargar Enviados`: Procesa y descarga todos los documentos de la bandeja de enviados.
    - `[3] 👤 Cambiar de usuario`: Permite cambiar de cuenta dentro de Quipux sin cerrar el bot.
    - `[0] 🚪 Salir`: Cierra la aplicación de forma segura.

4.  **Proceso de Descarga**:
    - El bot navegará página por página.
    - Creará una carpeta por cada documento.
    - Descargará el PDF principal y todos sus anexos (si los tiene).
    - Al finalizar cada bandeja, generará un archivo Excel con el resumen.

## 📂 Estructura de Archivos Generados

Los archivos se guardan en la carpeta `mis_respaldos/` con la siguiente estructura:

```text
mis_respaldos/
└── NOMBRE_USUARIO/
    ├── Recibidos/
    │   ├── quipux_recibidos.xlsx        <-- Reporte detallado
    │   ├── DOC_001_NUM_DOC/             <-- Carpeta por documento
    │   │   ├── documento.pdf            <-- Archivo principal
    │   │   └── anexos/                  <-- Carpeta de anexos
    │   │       ├── anexo1.pdf
    │   │       └── ...
    │   └── ...
    └── Enviados/
        ├── quipux_enviados.xlsx
        └── ...
```

## ⚠️ Notas Importantes

- **No cierres el navegador manualmente** a menos que el bot ya haya terminado o se lo indiques.
- Si el bot encuentra errores de red, intentará reintentar la descarga automáticamente.
- El archivo Excel incluye metadatos como remitente, asunto, fecha, número de documento y cantidad de anexos.

---
*Desarrollado para la automatización de respaldos en Quipux.*
