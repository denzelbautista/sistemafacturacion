from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import time
import os
from pathlib import Path

ruc_usuario = '20611634731'
user = '73273740'
ps = 'Hermanos3'

# 1. Configurar la ruta de descarga deseada
download_dir = Path(r'C:\Users\ysaba\Desktop\Ventas')  # Cambia esto a tu ruta
download_dir.mkdir(parents=True, exist_ok=True)  # Crea el directorio si no existe

os.environ['PLAYWRIGHT_DOWNLOAD_PATH'] = str(download_dir)

excel = pd.read_excel(r'C:\Users\ysaba\Desktop\Pedidos.xlsx', sheet_name='Detalles')

def error_auth(page, max_intentos=5, delay=2):
    for intento in range(max_intentos):
        try:
            # Verificar si hay un error de autenticación primero
            error_auth = page.query_selector('div#content:has-text("No se ha enviado correctamente los parametros de autenticacion")')
            if error_auth:
                print("⚠️ Error de autenticación detectado. Recargando...")
                page.click('li[data-id="11.5.3.1.1"]')  # Selector del botón
                time.sleep(delay)
                continue

            return True
                
        except Exception as e:
            print(f"Intento {intento + 1} fallido: {str(e)}")
            time.sleep(delay)
    
    print("❌ No se pudo cargar el formulario después de varios intentos")
    return False

# Scrapper sunat

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False) # Para ver el navegador

        # 2. Configurar el contexto del navegador para manejar descargas
        context = browser.new_context(accept_downloads=True)
        
        page = context.new_page()

        page.goto("https://e-menu.sunat.gob.pe/")
        page.wait_for_timeout(4000)

        page.fill('input[id="txtRuc"]', ruc_usuario)
        page.fill('input[id="txtUsuario"]', user)
        page.fill('input[id="txtContrasena"]', ps)
        page.click('button[id="btnAceptar"]')

        page.wait_for_load_state("networkidle")

        page.click('div[id="divOpcionServicio2"]')

        page.click('li[data-id="11"]') # Comprobantes de pago

        page.click('li[data-id="11.5"]') # SEE Sol

        page.click('li[data-id="11.5.3"]') # Factura electrónica

        page.click('li[data-id="11.5.3.1.1"]')  # Selector del botón

        page.wait_for_load_state("networkidle")

        # Aqui llega el iframe

        page.wait_for_selector('#iframeApplication')

        for index, row in excel.iterrows():

            page.wait_for_load_state("networkidle")

            # Aqui llega el iframe

            page.wait_for_selector('#iframeApplication')

            # ---------------------------------------------------------
            ruc = str(row['RUC'])
            empresa = str(row['EMPRESA'])
            precio_final = str(row['PRECIOFINAL'])
            pedido = str(row['PEDIDO'])
            cantidad = str(row['CANTIDAD'])
            destino = str(row['DESTINO'])
            descripcion = (f"{cantidad} {pedido} - {destino}").upper()
            # ---------------------------------------------------------

            frame = page.frame_locator('#iframeApplication')

            frame.locator('input[id="inicio.numeroDocumento"]').wait_for()

            frame.locator('input[id="inicio.numeroDocumento"]').fill(ruc) # Rellenar el ruc
            
            # Presionar afuera para salir del input
            frame.locator('div[id="widget_inicio.razonSocial"]').click()

            # Esperar a que la página procese el RUC (si es necesario)
            page.wait_for_timeout(3000)

            # Checkbox domicilio
            frame.locator('input[name="opcionDomicilioCliente"][value="1"]').click()

            # Verificar que quedó marcado correctamente (opcional)
            checkdomicilio = frame.locator('input#inicio\\.subTipoDC01')
            if checkdomicilio.get_attribute('aria-checked') != 'true':
                # Si no se marcó, intentar de nuevo con más tiempo
                print('⚠️ Primer checkbox no se ha validado aún')
                page.wait_for_timeout(2000)
                checkdomicilio.click()
            
            print('✅ Checkbox domicilio validado correctamente')

            page.wait_for_timeout(2000)

            frame.locator('span[widgetid="inicio.botonGrabarDocumento"]').click() # Aqui se accede a definir la factura
            
            # DEFINIENDO LA FACTURA

            frame.locator('span[id="factura.addItemButton"]').click() # Adicionamos el item

            frame.locator('input[name="tipoItem"][value="TI02"]').click() # Intentar clickear

            checkservicio = frame.locator('input#item\\.subTipoTI02')

            if checkservicio.get_attribute('aria-checked') != 'true':
                # Si no se marcó, intentar de nuevo con más tiempo
                print('⚠️ Segundo checkbox no se ha validado aún')
                page.wait_for_timeout(2000)
                checkservicio.click()
            
            print('✅ Checkbox tipo de servicio validado correctamente')

            frame.locator('textarea[widgetid="item.descripcion"]').fill(descripcion) # Descripción

            frame.locator('input[id="item.precioUnitario"]').click(click_count=3)  # Selecciona todo
            page.keyboard.press('Delete')
            
            frame.locator('input[id="item.precioUnitario"]').fill(precio_final) # Precio

            page.wait_for_timeout(3000)

            # Presionar afuera para salir del input
            frame.locator('div[id="widget_item.precioDescuento"]').click()
            frame.locator('input[name="tipoItem"][value="TI02"]').click() # Intentar clickear por siacaso

            page.wait_for_timeout(3000)

            frame.locator('span[id="item.botonAceptar"]').click() # Para salir de la lista de items

            frame.locator('span[id="factura.botonGrabarDocumento"]').click() # Continuar

            frame.locator('span[id="docsrel.botonGrabarDocumento"]').click() # Continuar

            frame.locator('span[id="factura-preliminar.botonGrabarDocumento"]').click() # Emitir factura

            frame.locator('span[id="dlgBtnAceptarConfirm"]').click() # Confirmar emitir factura

            # 3. Dentro de tu código existente:
            with page.expect_download() as download_info:
                frame.locator('span[id="dijit_form_Button_2"]').click()  # Click para descargar
                
            # 4. Obtener el objeto de descarga
            download = download_info.value
            fecha_formateada = datetime.now().strftime("%H-%M-%d-%m") + f"-{int(datetime.now().microsecond / 1000):03d}"
            # 5. Definir el nombre y ruta final del archivo
            nombre_personalizado = f"{empresa} {fecha_formateada}.pdf"  # Cambia esto al nombre que quieras
            ruta_final = download_dir / nombre_personalizado

            # 6. Mover y renombrar el archivo
            download.save_as(ruta_final)

            page.wait_for_timeout(3000) # Antes 35000 tiempo para iguales

            page.click('li[data-id="11.5.3.1.1"]')  # Volver para emitir nueva factura

            print(f"✅ Factura {nombre_personalizado} procesada con éxito.")

            # FIN

        print("Facturas procesadas")
        page.wait_for_timeout(3000)

        page.click('button[id="btnSalir"]')

        browser.close()

if __name__ == "__main__":
    main()
