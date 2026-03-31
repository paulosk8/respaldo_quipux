import asyncio
import os
import re
import sys
from datetime import datetime
from playwright.async_api import async_playwright

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("⚠ openpyxl no instalado. Ejecuta: pip install openpyxl")
    print("  El Excel no se generará, pero las descargas continuarán.\n")


# ─────────────────────────────────────────────────────────
# Generar archivo Excel a partir de datos de documentos
# ─────────────────────────────────────────────────────────
def generar_excel(all_docs_data, xlsx_path, sheet_name="Documentos"):
    if not HAS_OPENPYXL or not all_docs_data:
        return
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="006394", end_color="006394", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_font = Font(name="Calibri", size=10)
    cell_align = Alignment(vertical="top", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    even_fill = PatternFill(start_color="E8F4FD", end_color="E8F4FD", fill_type="solid")

    headers = ["No.", "De", "Asunto", "Fecha Documento",
               "Número Documento", "No. Referencia",
               "Usuario Anterior", "Carpeta Descarga", "Anexos"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    for row_idx, doc_data in enumerate(all_docs_data, 2):
        values = [
            doc_data["No"], doc_data["De"], doc_data["Asunto"],
            doc_data["Fecha Documento"], doc_data["Número Documento"],
            doc_data["No. Referencia"], doc_data["Usuario Anterior"],
            doc_data["Carpeta"], doc_data["Anexos"],
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = cell_font
            cell.alignment = cell_align
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = even_fill

    col_widths = [6, 35, 50, 18, 25, 18, 30, 30, 8]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(xlsx_path)
    print(f"\n  📊 Excel generado: {xlsx_path}")


# ─────────────────────────────────────────────────────────
# Descargar todos los documentos de la bandeja actual
# ─────────────────────────────────────────────────────────
async def descargar_bandeja(main_frame, context, download_dir, bandeja_nombre):
    """Descarga todos los documentos (con paginación) de la bandeja activa."""

    os.makedirs(download_dir, exist_ok=True)

    # Esperar a que se cargue el contenido
    try:
        await main_frame.wait_for_selector("tr.listado1, tr.listado2", timeout=15000)
    except Exception:
        print("       Esperando carga AJAX...")
        await asyncio.sleep(5)

    # Detectar total de páginas
    total_pages = await main_frame.evaluate("""
        () => {
            const text = document.body.innerText;
            const match = text.match(/Página\\s+(\\d+)\\/(\\d+)/);
            if (match) return parseInt(match[2]);
            return 1;
        }
    """)
    print(f"       Páginas encontradas: {total_pages}")

    descargados = 0
    errores = 0
    total_docs = 0
    all_docs_data = []

    for current_page in range(1, total_pages + 1):
        print(f"\n{'=' * 60}")
        print(f"  📄 PÁGINA {current_page} / {total_pages} — {bandeja_nombre}")
        print(f"{'=' * 60}")

        # Navegar a la página si no es la primera
        if current_page > 1:
            print(f"  Navegando a página {current_page}...")
            await main_frame.evaluate(f"""
                () => {{
                    paginador_reload_div('adodb_next_page={current_page}');
                }}
            """)
            await asyncio.sleep(3)
            try:
                await main_frame.wait_for_selector("tr.listado1, tr.listado2", timeout=15000)
            except Exception:
                await asyncio.sleep(3)

        # Extraer documentos de la página actual
        docs = await main_frame.evaluate("""
            () => {
                const docs = [];
                const rows = document.querySelectorAll('tr.listado1, tr.listado2');
                rows.forEach((row, idx) => {
                    const cells = row.querySelectorAll('td');
                    const links = row.querySelectorAll('a[href*="mostrar_documento"]');
                    let radicado = '', textrad = '', carpeta = '';
                    for (const link of links) {
                        const href = link.getAttribute('href') || '';
                        const match = href.match(/mostrar_documento\\("([^"]+)","([^"]+)","([^"]+)"\\)/);
                        if (match) {
                            radicado = match[1];
                            textrad = match[2];
                            carpeta = match[3];
                            break;
                        }
                    }
                    if (!radicado) {
                        const allLinks = row.querySelectorAll('a');
                        for (const link of allLinks) {
                            const onclick = link.getAttribute('onclick') || '';
                            const href = link.getAttribute('href') || '';
                            const text = onclick + ' ' + href;
                            const match = text.match(/mostrar_documento\\(['"](\\d+)['"],\\s*['"]([^'"]+)['"],\\s*['"]([^'"]+)['"]\\)/);
                            if (match) {
                                radicado = match[1];
                                textrad = match[2];
                                carpeta = match[3];
                                break;
                            }
                        }
                    }

                    let asunto = '', remitente = '', fecha = '', numDoc = '';
                    let noReferencia = '', usuarioAnterior = '';
                    if (cells.length >= 8) {
                        remitente = (cells[5]?.innerText || '').trim();
                        asunto = (cells[6]?.innerText || '').trim();
                        fecha = (cells[7]?.innerText || '').trim();
                        numDoc = (cells[8]?.innerText || '').trim();
                    }
                    if (cells.length >= 10) {
                        noReferencia = (cells[9]?.innerText || '').trim();
                    }
                    if (cells.length >= 11) {
                        usuarioAnterior = (cells[10]?.innerText || '').trim();
                    }

                    if (radicado) {
                        docs.push({
                            index: idx, radicado, textrad, carpeta,
                            asunto, remitente, fecha, numDoc,
                            noReferencia, usuarioAnterior
                        });
                    }
                });
                return docs;
            }
        """)

        print(f"  Documentos en esta página: {len(docs)}")

        if not docs:
            print("  Sin documentos, saltando...")
            continue

        for doc in docs:
            total_docs += 1
            radicado = doc["radicado"]
            textrad = doc["textrad"]
            carpeta = doc["carpeta"]
            asunto = doc["asunto"][:80]
            numDoc = doc["numDoc"]
            noReferencia = doc.get("noReferencia", "")
            usuarioAnterior = doc.get("usuarioAnterior", "")

            safe_num = re.sub(r'[^\w\-]', '_', numDoc or textrad)
            doc_folder = f"DOC_{total_docs:03d}_{safe_num}"
            doc_dir = os.path.join(download_dir, doc_folder)
            anexos_dir = os.path.join(doc_dir, "anexos")
            os.makedirs(doc_dir, exist_ok=True)

            print(f"  ─── [{total_docs}] {numDoc} ───")
            print(f"  Asunto: {asunto}")

            try:
                doc_url = (
                    f"https://quipux.espe.edu.ec/verradicado.php"
                    f"?verrad={radicado}&textrad={textrad}"
                    f"&carpeta={carpeta}&menu_ver_tmp=3&tipo_ventana=popup"
                )
                doc_page = await context.new_page()
                await doc_page.goto(doc_url, wait_until="domcontentloaded", timeout=60000)
                await asyncio.sleep(3)

                # ---- DOCUMENTO PRINCIPAL ----
                doc_downloaded = False

                ver_doc_selectors = [
                    'a:has-text("Ver Documento")',
                    'a:has-text("Ver documento")',
                    'a:has-text("ver documento")',
                    'a[href*="ver_documento"]',
                    'a[href*="anexos_descargar_archivo"]',
                ]
                for sel in ver_doc_selectors:
                    try:
                        ver_link = await doc_page.query_selector(sel)
                        if ver_link:
                            try:
                                async with doc_page.expect_download(timeout=15000) as dl:
                                    await ver_link.click()
                                download = await dl.value
                                fname = download.suggested_filename or f"documento_{total_docs}.pdf"
                                await download.save_as(os.path.join(doc_dir, fname))
                                print(f"  ✓ Documento: {fname}")
                                doc_downloaded = True
                                descargados += 1
                            except Exception:
                                pass
                        if doc_downloaded:
                            break
                    except Exception:
                        continue

                if not doc_downloaded:
                    try:
                        iframe_src = await doc_page.evaluate("""
                            () => {
                                const ifr = document.getElementById('ifr_descargar_archivo');
                                return ifr ? ifr.src : null;
                            }
                        """)
                        if iframe_src and 'radi_nume' in iframe_src:
                            dl_page = await context.new_page()
                            try:
                                async with dl_page.expect_download(timeout=15000) as dl:
                                    await dl_page.goto(iframe_src)
                                download = await dl.value
                                fname = download.suggested_filename or f"documento_{total_docs}.pdf"
                                await download.save_as(os.path.join(doc_dir, fname))
                                print(f"  ✓ Documento (iframe): {fname}")
                                doc_downloaded = True
                                descargados += 1
                            except Exception:
                                pass
                            finally:
                                await dl_page.close()
                    except Exception:
                        pass

                if not doc_downloaded:
                    try:
                        async with doc_page.expect_download(timeout=15000) as dl:
                            await doc_page.evaluate("""
                                () => {
                                    if (typeof vista_previa === 'function') vista_previa();
                                    else if (typeof fjs_radicado_descargar_archivo === 'function')
                                        fjs_radicado_descargar_archivo(radi_nume, '', 0, 'download');
                                }
                            """)
                        download = await dl.value
                        fname = download.suggested_filename or f"documento_{total_docs}.pdf"
                        await download.save_as(os.path.join(doc_dir, fname))
                        print(f"  ✓ Documento (JS): {fname}")
                        doc_downloaded = True
                        descargados += 1
                    except Exception:
                        pass

                if not doc_downloaded:
                    print(f"  ⚠ No se pudo descargar el documento principal.")

                # ---- ANEXOS ----
                print(f"  Buscando anexos...")
                try:
                    anexos_url = (
                        f"https://quipux.espe.edu.ec/verradicado.php"
                        f"?carpeta={carpeta}&verrad={radicado}"
                        f"&textrad={textrad}&estadisticas=0"
                        f"&verPDF=1&irVerRad=1&tipo_ventana=popup&menu_ver=2"
                    )
                    await doc_page.goto(anexos_url, wait_until="domcontentloaded", timeout=60000)
                    await asyncio.sleep(2)

                    div_status = 'loading'
                    for _ in range(12):
                        div_status = await doc_page.evaluate("""
                            () => {
                                const div = document.getElementById('div_anexos_lista_archivos');
                                if (!div) return 'no_div';
                                const html = div.innerHTML.trim();
                                if (html === '') return 'empty';
                                if (html.includes('no tiene archivos anexos')) return 'sin_anexos';
                                if (html.includes('anexos_descargar_archivo')) return 'tiene_anexos';
                                if (html.length > 50) return 'loaded';
                                return 'loading';
                            }
                        """)
                        if div_status in ('sin_anexos', 'tiene_anexos', 'loaded', 'no_div'):
                            break
                        await asyncio.sleep(0.5)

                    if div_status == 'sin_anexos':
                        print(f"  (Sin anexos)")
                    elif div_status == 'no_div':
                        print(f"  ⚠ div_anexos no encontrado")
                    else:
                        anexo_data = await doc_page.evaluate("""
                            () => {
                                const div = document.getElementById('div_anexos_lista_archivos');
                                if (!div) return [];
                                const html = div.innerHTML;
                                const results = [];
                                const fnName = "anexos_descargar_archivo(";
                                let startIdx = 0;
                                while (true) {
                                    const pos = html.indexOf(fnName, startIdx);
                                    if (pos === -1) break;
                                    startIdx = pos + fnName.length;
                                    const closePos = html.indexOf(')', startIdx);
                                    if (closePos === -1) break;
                                    const argsStr = html.substring(startIdx, closePos);
                                    const args = argsStr.split(',').map(a =>
                                        a.trim().replace(/^['"]/, '').replace(/['"]$/, '').trim()
                                    );
                                    if (args.length >= 3) {
                                        results.push({
                                            radicado: args[0],
                                            anex_codigo: args[1],
                                            arch_tipo: args[2],
                                            tipo_descarga: args.length >= 4 ? args[3] : 'download'
                                        });
                                    }
                                }
                                return results;
                            }
                        """)

                        seen = set()
                        unique_anexos = []
                        for a in anexo_data:
                            key = f"{a['radicado']}_{a['anex_codigo']}_{a['arch_tipo']}"
                            td = a.get('tipo_descarga', '').strip("'\"")
                            if key not in seen and 'embeded' not in td:
                                seen.add(key)
                                unique_anexos.append(a)

                        if unique_anexos:
                            os.makedirs(anexos_dir, exist_ok=True)
                            print(f"  Anexos: {len(unique_anexos)}")
                            for j, anexo in enumerate(unique_anexos):
                                try:
                                    rad = anexo['radicado']
                                    anex = anexo['anex_codigo']
                                    arch = anexo['arch_tipo']
                                    async with doc_page.expect_download(timeout=30000) as dl:
                                        await doc_page.evaluate(f"""
                                            () => {{
                                                anexos_descargar_archivo('{rad}', '{anex}', {arch});
                                            }}
                                        """)
                                    download = await dl.value
                                    fname = download.suggested_filename or f"anexo_{j + 1}.bin"
                                    await download.save_as(os.path.join(anexos_dir, fname))
                                    print(f"  ✓ Anexo: {fname}")
                                    await asyncio.sleep(1)
                                except Exception as e:
                                    print(f"  ⚠ Error anexo {j + 1}: {e}")
                        else:
                            print(f"  (Sin anexos)")

                except Exception as e:
                    import traceback
                    print(f"  ⚠ Error buscando anexos: {e}")
                    traceback.print_exc()

                await doc_page.close()

            except Exception as e:
                errores += 1
                print(f"  ✗ Error general: {e}")
                while len(context.pages) > 1:
                    try:
                        await context.pages[-1].close()
                    except Exception:
                        break

            # Datos para Excel
            num_anexos_descargados = 0
            anexos_path_check = os.path.join(doc_dir, "anexos")
            if os.path.isdir(anexos_path_check):
                num_anexos_descargados = len(os.listdir(anexos_path_check))

            all_docs_data.append({
                "No": total_docs,
                "De": doc.get("remitente", ""),
                "Asunto": doc.get("asunto", ""),
                "Fecha Documento": doc.get("fecha", ""),
                "Número Documento": numDoc,
                "No. Referencia": noReferencia,
                "Usuario Anterior": usuarioAnterior,
                "Carpeta": doc_folder,
                "Anexos": num_anexos_descargados,
            })
            print()

    # Generar Excel
    if all_docs_data:
        xlsx_name = f"quipux_{bandeja_nombre.lower().replace(' ', '_')}.xlsx"
        xlsx_path = os.path.join(download_dir, xlsx_name)
        generar_excel(all_docs_data, xlsx_path, bandeja_nombre)

    # Resumen
    print("\n" + "=" * 60)
    print(f"  RESUMEN — {bandeja_nombre}")
    print(f"  Total páginas:        {total_pages}")
    print(f"  Total documentos:     {total_docs}")
    print(f"  Descargados:          {descargados}")
    print(f"  Errores:              {errores}")
    print(f"  Carpeta:              {download_dir}")
    print("=" * 60)

    # Listar carpetas
    folders = sorted([
        d for d in os.listdir(download_dir)
        if os.path.isdir(os.path.join(download_dir, d)) and d.startswith("DOC_")
    ])
    if folders:
        print(f"\n  📁 Carpetas ({len(folders)}):")
        for folder in folders:
            full_path = os.path.join(download_dir, folder)
            main_files = [f for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f))]
            print(f"    📁 {folder}/ ({len(main_files)} doc)")
            anexos_path = os.path.join(full_path, "anexos")
            if os.path.isdir(anexos_path):
                anexo_files = os.listdir(anexos_path)
                if anexo_files:
                    print(f"       📁 anexos/ ({len(anexo_files)} archivos)")

    return total_docs, descargados, errores


# ─────────────────────────────────────────────────────────
# Programa principal con menú interactivo
# ─────────────────────────────────────────────────────────
async def run():
    async with async_playwright() as p:
        BASE_DIR = os.path.abspath(os.getenv("DOWNLOAD_DIR", "./mis_respaldos"))
        os.makedirs(BASE_DIR, exist_ok=True)

        print("=" * 60)
        print("  QUIPUX - Bot de descarga de documentos")
        print("=" * 60)

        browser = await p.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors"],
        )
        context = await browser.new_context(
            accept_downloads=True,
            ignore_https_errors=True,
        )
        await context.add_init_script("""
            Object.defineProperty(window, 'menubar', {
                get: () => ({ visible: false })
            });
            Object.defineProperty(window, 'toolbar', {
                get: () => ({ visible: false })
            });
        """)

        page = await context.new_page()
        page.set_default_timeout(60000)

        print("\n[1/3] Abriendo login.php...")
        await page.goto("https://quipux.espe.edu.ec/login.php", wait_until="domcontentloaded")
        await asyncio.sleep(2)
        print("       Ingresa tu usuario y clave en el navegador.")
        input("\n>>> Presiona ENTER cuando veas la bandeja de documentos... ")

        # ─── Encontrar página y frame ───
        print("\n[2/3] Buscando la bandeja de documentos...")
        target_page = None
        for pg in context.pages:
            try:
                url = pg.url
                if "index_frames" in url or "cuerpo" in url:
                    target_page = pg
            except Exception:
                pass
        if not target_page:
            for pg in context.pages:
                try:
                    _ = pg.url
                    target_page = pg
                except Exception:
                    pass
        if not target_page:
            print("ERROR: No se encontró ventana activa.")
            await browser.close()
            sys.exit(1)
        print(f"       ✓ Ventana: {target_page.url}")

        print("\n[3/3] Buscando frame principal...")
        await asyncio.sleep(2)

        # ─── Función para buscar mainFrame ───
        async def obtener_main_frame():
            for f in target_page.frames:
                if f.name == "mainFrame":
                    return f
            for f in target_page.frames:
                if f.url and "about:blank" not in f.url and f != target_page.main_frame:
                    try:
                        fc = await f.content()
                        if "mostrar_documento" in fc or "div_cuerpo" in fc:
                            return f
                    except Exception:
                        pass
            content = await target_page.content()
            if "mostrar_documento" in content or "div_cuerpo" in content:
                return target_page.main_frame
            return None

        # ─── Función para leer usuario actual ───
        async def obtener_usuario_actual():
            """Lee el nombre del usuario actual buscando en todos los frames."""
            nombre = None
            for frame in target_page.frames:
                try:
                    nombre = await frame.evaluate("""
                        () => {
                            const tds = document.querySelectorAll('td');
                            for (const td of tds) {
                                const text = td.innerText || '';
                                if (text.includes('Usuario actual:')) {
                                    const parts = text.split('Usuario actual:');
                                    if (parts[1]) return parts[1].trim().split('\\n')[0].trim();
                                }
                            }
                            const body = document.body?.innerText || '';
                            const m = body.match(/Usuario actual:\\s*(.+)/);
                            if (m) return m[1].trim().split('\\n')[0].trim();
                            return null;
                        }
                    """)
                    if nombre:
                        break
                except Exception:
                    continue
            return nombre or "Usuario desconocido"

        # ══════════════════════════════════════════════
        # SELECCIÓN INICIAL DE USUARIO
        # ══════════════════════════════════════════════
        print("\n" + "═" * 60)
        print("  Si necesitas cambiar de usuario en Quipux,")
        print("  hazlo ahora en el navegador.")
        print("═" * 60)
        input("\n>>> Presiona ENTER cuando estés listo con el usuario deseado... ")

        usuario_actual = await obtener_usuario_actual()

        # ══════════════════════════════════════════════
        # MENÚ INTERACTIVO
        # ══════════════════════════════════════════════
        while True:
            print("\n" + "═" * 60)
            print("  ╔══════════════════════════════════════════════════╗")
            print("  ║          QUIPUX — Menú Principal                ║")
            print("  ╚══════════════════════════════════════════════════╝")
            print(f"\n  👤 Usuario: {usuario_actual}")
            print("\n  ─── Opciones ───")
            print("  [1] 📥 Descargar Recibidos")
            print("  [2] 📤 Descargar Enviados")
            print("  [3] 👤 Cambiar de usuario")
            print("  [0] 🚪 Salir")
            print("═" * 60)

            opcion = input("\n  Selecciona una opción: ").strip()

            if opcion == "0":
                print("\n  👋 ¡Hasta luego!")
                break

            elif opcion in ("1", "2"):
                if opcion == "1":
                    bandeja_nombre = "Recibidos"
                    carpeta_codigo = "2"
                else:
                    bandeja_nombre = "Enviados"
                    carpeta_codigo = "8"

                print(f"\n{'=' * 60}")
                emoji = "📥" if opcion == "1" else "📤"
                print(f"  {emoji} DESCARGA DE {bandeja_nombre.upper()}")
                print(f"  👤 {usuario_actual}")
                print(f"{'=' * 60}")

                # Navegar a la bandeja seleccionada
                main_frame = await obtener_main_frame()
                if main_frame:
                    try:
                        await main_frame.evaluate(f"""
                            () => {{
                                if (typeof llamarListado === 'function') {{
                                    llamarListado('{bandeja_nombre}', '{carpeta_codigo}');
                                }} else {{
                                    location.href = 'cuerpo.php?nomcarpeta={bandeja_nombre}&carpeta={carpeta_codigo}&adodb_next_page=1';
                                }}
                            }}
                        """)
                        await asyncio.sleep(4)
                    except Exception as e:
                        pass

                # Obtener mainFrame nuevamente porque al navegar se desmonta (Frame detached)
                main_frame = await obtener_main_frame()
                if not main_frame:
                    print(f"  ⚠ Error: No se encontró el frame principal para descargar {bandeja_nombre}.")
                    continue

                safe_user = re.sub(r'[^\w]', '_', usuario_actual)[:30]
                dl_dir = os.path.join(BASE_DIR, safe_user, bandeja_nombre)
                await descargar_bandeja(main_frame, context, dl_dir, f"{bandeja_nombre} — {usuario_actual}")

                print("\n" + "─" * 60)
                input("  Presiona ENTER para volver al menú... ")

            elif opcion == "3":
                # ── Cambiar usuario ──
                print("\n" + "─" * 60)
                print("  👤 CAMBIAR DE USUARIO")
                print("─" * 60)
                print("\n  Ve al navegador y selecciona el usuario deseado.")
                print("  (Puedes hacerlo desde el panel izquierdo de Quipux)")
                input("\n>>> Presiona ENTER cuando hayas cambiado de usuario... ")

                await asyncio.sleep(2)
                usuario_actual = await obtener_usuario_actual()
                print(f"\n  ✓ Usuario detectado: {usuario_actual}")

            else:
                print("  ⚠ Opción no válida")

        await browser.close()


asyncio.run(run())