import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from io import BytesIO
import os
import re

# Ruta base donde se encuentran las carpetas mensuales
RUTA_BASE = "C:\\Users\\USER\\SRI\\LAYHER"

# Función para extraer el XML dentro de CDATA y procesarlo
def extraer_xml_cdata(xml_string):
    # Busca el contenido dentro de <![CDATA[...]]>
    match = re.search(r'<!\[CDATA\[(.*?)\]\]>', xml_string, re.DOTALL)
    if match:
        return match.group(1)
    return None

# Función para leer los archivos XML desde una carpeta específica y extraer los datos
def leer_facturas_desde_carpeta(carpeta):
    datos_facturas = []

    # Iteramos sobre los archivos XML en la carpeta
    for archivo in os.listdir(carpeta):
        if archivo.endswith('.xml'):
            archivo_xml = os.path.join(carpeta, archivo)
            with open(archivo_xml, 'r', encoding='utf-8') as file:
                contenido_xml = file.read()
                xml_interno = extraer_xml_cdata(contenido_xml)
                
                if xml_interno:
                    # Parseamos el XML interno de la factura
                    root = ET.fromstring(xml_interno)
                    
                    # Extraemos el número de factura a partir de estab, ptoEmi, y secuencial
                    info_tributaria = root.find('.//infoTributaria')
                    if info_tributaria is not None:
                        razon_social = info_tributaria.find('razonSocial').text if info_tributaria.find('razonSocial') is not None else "N/A"
                        ruc = info_tributaria.find('ruc').text if info_tributaria.find('ruc') is not None else "N/A"
                        establecimiento = info_tributaria.find('estab').text if info_tributaria.find('estab') is not None else "N/A"
                        punto_emision = info_tributaria.find('ptoEmi').text if info_tributaria.find('ptoEmi') is not None else "N/A"
                        secuencial = info_tributaria.find('secuencial').text if info_tributaria.find('secuencial') is not None else "N/A"
                        numero_factura = f"{establecimiento}-{punto_emision}-{secuencial}"
                    else:
                        numero_factura = "N/A"
                        razon_social = "N/A"
                        ruc = "N/A"

                    # Extraer fecha de emisión de la factura
                    info_factura = root.find('.//infoFactura')
                    if info_factura is not None:
                        fecha_emision = info_factura.find('fechaEmision').text if info_factura.find('fechaEmision') is not None else "N/A"
                        subtotal_15 = 0.00
                        subtotal_0 = 0.00
                        subtotal_sin_impuestos = float(info_factura.find('totalSinImpuestos').text) if info_factura.find('totalSinImpuestos') is not None else 0.00
                        total_descuento = float(info_factura.find('totalDescuento').text) if info_factura.find('totalDescuento') is not None else 0.00
                        valor_total = float(info_factura.find('importeTotal').text) if info_factura.find('importeTotal') is not None else 0.00
                        propina = float(info_factura.find('propina').text) if info_factura.find('propina') is not None else 0.00
                        iva_15 = 0.00
                        ice = 0.00
                        irbpnr = 0.00
                        productos_con_iva = []
                        productos_sin_iva = []
                        retencion = 0.00
                        
                        # Revisamos los impuestos por cada producto
                        for impuesto in info_factura.findall('.//totalImpuesto'):
                            codigo_porcentaje = impuesto.find('codigoPorcentaje').text if impuesto.find('codigoPorcentaje') is not None else ""
                            base_imponible = float(impuesto.find('baseImponible').text) if impuesto.find('baseImponible') is not None else 0.00
                            valor_impuesto = float(impuesto.find('valor').text) if impuesto.find('valor').text is not None else 0.00
                            
                            if codigo_porcentaje == "4":  # Asumiendo que el código "4" es IVA 15%
                                subtotal_15 += base_imponible
                                iva_15 += valor_impuesto
                            elif codigo_porcentaje == "0":  # Asumiendo que el código "0" es para IVA 0%
                                subtotal_0 += base_imponible
                            elif codigo_porcentaje == "3":  # Retención de impuestos
                                retencion += valor_impuesto
                            elif codigo_porcentaje == "6":  # IRBPNR
                                irbpnr += valor_impuesto
                            
                        # Extraemos los detalles de los productos
                        for detalle in root.findall('.//detalle'):
                            descripcion = detalle.find('descripcion').text if detalle.find('descripcion') is not None else "N/A"
                            cantidad = detalle.find('cantidad').text if detalle.find('cantidad') is not None else "N/A"
                            precio_unitario = detalle.find('precioUnitario').text if detalle.find('precioUnitario') is not None else "N/A"
                            precio_total = detalle.find('precioTotalSinImpuesto').text if detalle.find('precioTotalSinImpuesto') is not None else "N/A"
                            
                            impuestos = detalle.find('impuestos')
                            tiene_iva = False
                            if impuestos is not None:
                                for impuesto in impuestos.findall('impuesto'):
                                    tarifa = impuesto.find('tarifa').text if impuesto.find('tarifa') is not None else "N/A"
                                    if tarifa != "0":  # Tiene IVA
                                        tiene_iva = True
                                        productos_con_iva.append(descripcion)
                                    else:
                                        productos_sin_iva.append(descripcion)
                            
                            # Almacenar los datos del producto
                            datos_facturas.append({
                                'Número de Factura': numero_factura,
                                'Fecha de Emisión': fecha_emision,
                                'Razón Social': razon_social,
                                'RUC': ruc,
                                'Descripción Producto': descripcion,
                                'Cantidad': cantidad,
                                'Precio Unitario': precio_unitario,
                                'Precio Total': precio_total,
                                'Tiene IVA': 'Sí' if tiene_iva else 'No'
                            })
                        
                        # Almacenar los totales y datos de la factura
                        datos_facturas.append({
                            'Número de Factura': numero_factura,
                            'Fecha de Emisión': fecha_emision,
                            'Razón Social': razon_social,
                            'RUC': ruc,
                            'Subtotal 15%': subtotal_15,
                            'Subtotal 0%': subtotal_0,
                            'Subtotal Sin Impuestos': subtotal_sin_impuestos,
                            'Total Descuento': total_descuento,
                            'IVA 15%': iva_15,
                            'ICE': ice,
                            'IRBPNR': irbpnr,
                            'Propina': propina,
                            'Retención': retencion,
                            'Valor Total': valor_total
                        })
    
    # Convertimos la lista de productos y facturas en un DataFrame de pandas
    df_facturas = pd.DataFrame(datos_facturas)
    
    return df_facturas

# Función para generar el archivo Excel en memoria
def generar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Interfaz de usuario con Streamlit
st.title("Generador de Excel desde Facturas Electrónicas XML")

# Pedimos al usuario que ingrese el nombre del mes en mayúsculas
mes = st.text_input("Introduce el mes (por ejemplo: SEPTIEMBRE):").strip()

if mes:
    # Creamos la ruta de la carpeta del mes basado en la entrada del usuario
    carpeta_facturas = os.path.join(RUTA_BASE, mes)

    # Verificamos si la carpeta existe
    if os.path.exists(carpeta_facturas):
        st.write(f"Procesando los archivos XML en la carpeta: {carpeta_facturas}")
        facturas_df = leer_facturas_desde_carpeta(carpeta_facturas)

        # Verificamos si se extrajeron datos
        if not facturas_df.empty:
            # Mostrar los datos extraídos
            st.write("Datos extraídos de las facturas:")
            st.dataframe(facturas_df)

            # Generar archivo Excel
            excel_data = generar_excel(facturas_df)

            # Botón de descarga para el archivo Excel
            st.download_button(
                label="Descargar Excel",
                data=excel_data,
                file_name="facturas_electronicas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.write("No se encontraron datos en los archivos XML.")
    else:
        st.write(f"La carpeta {carpeta_facturas} no existe.")
