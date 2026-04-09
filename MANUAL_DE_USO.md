# Manual de Usuario - Bot de Respaldo Quipux

Este bot automatiza la descarga de documentos y anexos desde el sistema Quipux de la ESPE, organizándolos por carpetas y generando un reporte en Excel.

## 🛠️ Requisitos Previos

Asegúrate de tener instalado lo siguiente en tu computadora:

1.  **Python 3.8 o superior**: Puedes descargarlo desde [python.org](https://www.python.org/).
2.  **Bibliotecas de Python**: Instaladas mediante `pip`.
3.  **Navegador Chromium**: Instalado a través de Playwright.

## 🚀 Instalación

Sigue estos pasos para configurar el entorno:

### En macOS / Linux

1.  **Instalar las dependencias de Python**:
    Abre una terminal y ejecuta:
    ```bash
    pip install playwright openpyxl
    ```

2.  **Instalar el navegador necesario**:
    ```bash
    playwright install chromium
    ```

### En Windows

> ⚠️ **Importante**: En Windows es necesario que Python y pip estén correctamente configurados en el **PATH** del sistema. Si al ejecutar `pip` o `python3` aparece un error de "comando no reconocido", sigue estos pasos:

1.  **Configurar el PATH de Python**:
    - Abre el **menú Inicio** y busca **"Variables de entorno"**.
    - Haz clic en **"Editar las variables de entorno del sistema"**.
    - En la ventana de Propiedades del sistema, haz clic en **"Variables de entorno..."**.
    - En **"Variables del sistema"**, busca la variable `Path` y haz clic en **"Editar"**.
    - Agrega las siguientes rutas (ajusta según tu instalación de Python):
      ```
      C:\Users\TU_USUARIO\AppData\Local\Programs\Python\Python3XX\
      C:\Users\TU_USUARIO\AppData\Local\Programs\Python\Python3XX\Scripts\
      ```
    - Haz clic en **"Aceptar"** en todas las ventanas.
    - **Cierra y vuelve a abrir** la terminal (CMD o PowerShell) para que los cambios surtan efecto.

    > 💡 **Alternativa rápida**: Si reinstalás Python, aseguráte de marcar la casilla **"Add Python to PATH"** durante la instalación.

2.  **Instalar las dependencias de Python**:
    Abre una terminal (CMD o PowerShell) y ejecuta:
    ```bash
    pip install playwright openpyxl
    ```

3.  **Instalar el navegador necesario**:
    ```bash
    playwright install chromium
    ```

## 📋 Uso de la Aplicación

1.  **Ejecutar el Bot**:
    En la terminal, navega hasta la carpeta del proyecto y ejecuta:
    ```bash
    python3 bot_quipux.py
    ```

2.  **Inicio de Sesión (Manual)**:
    - Se abrirá una ventana del navegador de forma automática.
    - Ingresa tu usuario y contraseña de Quipux.
    - Una vez que veas la bandeja de entrada (Recibidos), regresa a la terminal y presiona **ENTER**.

3.  **Nombre de Carpeta**:
    - Al seleccionar el usuario, el bot te pedirá un **nombre de carpeta** donde se guardarán las descargas.
    - Puedes escribir un nombre personalizado o presionar **ENTER** para usar el nombre del usuario como carpeta por defecto.
    - Dentro de esta carpeta se crearán automáticamente subcarpetas `Recibidos/` y `Enviados/`.

4.  **Menú de Opciones**:
    El bot presentará un menú interactivo en la terminal:
    - `[1] 📥 Descargar Recibidos`: Procesa y descarga documentos de recibidos.
    - `[2] 📤 Descargar Enviados`: Procesa y descarga documentos de enviados.
    - `[3] 🗃️ Descargar Archivados`: Descarga documentos guardados en Archivados.
    - `[4] 🗂️ Carpetas Virtuales`: Navega interactivamente por el árbol de carpetas.
    - `[5] 👤 Cambiar de usuario`: Permite cambiar de cuenta y asignar nueva carpeta.
    - `[0] 🚪 Salir`: Cierra la aplicación de forma segura.

5.  **Filtrado por Fechas y Navegación**:
    - **Bandejas Normales (Opción 1 a 3)**: El bot te preguntará el **rango exacto de fechas** (`Fecha DESDE` y `Fecha HASTA`) que deseas extraer mediante la terminal.
    - **Carpetas Virtuales (Opción 4)**: 
        1. Al elegirla, seleccionarás un año del sistema superior en la web.
        2. Si la web cargó las opciones, el bot listará **agrupadamente** en la terminal las subcarpetas disponibles para que escojas por número. (O alternativamente, te pedirá que expulses el árbol dando un par de clics a los íconos de "+").
        3. El bot seleccionará la carpeta automáticamente vía JS. Al cargar los resultados, utilizarás las ***casillas de filtro nativo web de Quipux*** ('Fecha Desde / Fecha Hasta') presionado 'Buscar Documentos' localmente en la pantalla, y solo entonces le darás ENTER en la consola a que el bot recoja lo filtrado y extraído.

6.  **Proceso de Descarga**:
    - El bot navegará página por página.
    - Creará una carpeta por cada documento.
    - Descargará el PDF principal y todos sus anexos (si los tiene).
    - Al finalizar cada bandeja, generará un archivo Excel con el resumen.

## 📂 Estructura de Archivos Generados

Los archivos se guardan en la carpeta `mis_respaldos/` con la siguiente estructura:

```text
mis_respaldos/
└── NOMBRE_CARPETA/
    ├── Recibidos/
    │   ├── quipux_recibidos.xlsx        <-- Reporte detallado
    │   ├── DOC_001_NUM_DOC/             <-- Carpeta por documento
    │   │   ├── documento.pdf            <-- Archivo principal
    │   │   └── anexos/                  <-- Carpeta de anexos
    │   │       ├── anexo1.pdf
    │   │       └── ...
    │   └── ...
    ├── Enviados/
    │   ├── quipux_enviados.xlsx
    │   └── ...
    ├── Archivados/
    │   ├── quipux_archivados.xlsx
    │   └── ...
    └── Carpetas Virtuales/
        └── AÑO_NombreDeCarpeta/
            ├── reporte_virtual.xlsx
            └── ...
```

## ⚠️ Notas Importantes

- **No cierres el navegador manualmente** a menos que el bot ya haya terminado o se lo indiques.
- Si el bot encuentra errores de red, intentará reintentar la descarga automáticamente.
- El archivo Excel incluye metadatos como remitente, asunto, fecha, número de documento y cantidad de anexos.
- Al **cambiar de usuario**, el bot te pedirá nuevamente un nombre de carpeta para organizar las descargas del nuevo usuario.

---
*Desarrollado para la automatización de respaldos en Quipux.*
