import telnetlib3 as telnetlib
import asyncio
import datetime
import pandas as pd

async def enviar_codigo_por_telnet(ip_impresora, puerto, codigo):
    """
    Envía un código a una impresora a través de Telnet.

    Args:
        ip_impresora (str): La dirección IP de la impresora.
        puerto (int): El puerto Telnet de la impresora
        codigo (str): El código a enviar a la impresora.
    """
    try:
        string_send = str(codigo)
        reader, writer = await telnetlib.open_connection(ip_impresora, puerto)
        print(f"Enviando código ZPL ({len(string_send)} bytes)...")
        writer.write(string_send)
        await writer.drain()
        
        writer.close()
        
        print(f"Código enviado exitosamente a {ip_impresora}:{puerto}")
    except ConnectionRefusedError:
        print(f"Error: Conexión rechazada a {ip_impresora}:{puerto}. Asegúrese de que la impresora esté encendida y el servicio Telnet esté habilitado.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")
        
async def imprimir_desde_excel(ruta_excel, ip_impresora, puerto_impresora, zpl_template):
    """
    Lee datos de un archivo Excel y envía una etiqueta ZPL por cada fila.

    Args:
        ruta_excel (str): Ruta completa al archivo Excel (.xlsx).
        ip_impresora (str): La dirección IP de la impresora.
        puerto_impresora (int): El puerto Telnet/TCP de la impresora.
        zpl_template (str): La plantilla ZPL con marcadores de posición.
    """
    print(f"Cargando datos desde: {ruta_excel}")
    try:
        df = pd.read_excel(ruta_excel)
        print(f"Se encontraron {len(df)} entradas en el archivo Excel.")
    except FileNotFoundError:
        print(f"Error: El archivo Excel no se encontró en la ruta: {ruta_excel}")
        return
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return

    # Iterar sobre cada fila del DataFrame
    # df.iterrows() devuelve un índice y una Serie (que se puede convertir a diccionario)
    for index, row in df.iterrows():
        print(f"\n--- Procesando entrada #{index + 1} ---")
        
        # Convertir la fila a un diccionario para usar con .format()
        # .astype(str) es crucial para asegurar que todos los valores sean cadenas
        # antes de intentar formatear la plantilla ZPL. Esto previene errores si
        # hay números o fechas en Excel que no se manejen bien directamente.
        datos_fila = row.astype(str).to_dict()

        # Si la columna 'Fecha' existe y es un objeto fecha, la formateamos.
        if 'Fecha' in row and isinstance(row['Fecha'], (pd.Timestamp, datetime.date)):
             datos_fila['Fecha'] = row['Fecha'].strftime("%d/%m/%Y")
        elif 'Fecha' in row and isinstance(row['Fecha'], str):
             # Si ya es string, la usamos tal cual
             pass
        else:
             # Si no hay fecha o es de un tipo inesperado, puedes poner un valor por defecto o lanzar un error
             datos_fila['Fecha'] = "Fecha no disponible" # O un valor vacío ""
             
        print(f"Datos de la fila {index + 1}:")
        print(datos_fila)
        print("--------------------------------------------------")

        try:
            # Formatear el ZPL con los datos de la fila actual
            # Usamos **datos_fila para desempaquetar el diccionario como argumentos con nombre
            codigo_zpl_final = zpl_template.format(**datos_fila)
            print(f"ZPL generado para la entrada #{index + 1}.")
            # print(codigo_zpl_final) # Descomenta para ver el ZPL generado

            # Enviar el ZPL a la impresora y obtener el estado
            respuesta_estado = await enviar_codigo_por_telnet(
                ip_impresora, puerto_impresora, codigo_zpl_final
            )
            if respuesta_estado:
                print(f"Etiqueta #{index + 1} enviada y estado recibido.")

            # Pequeña pausa entre etiquetas para no saturar la impresora
            await asyncio.sleep(1) 

        except KeyError as ke:
            print(f"Error: Marcador de posición '{ke}' no encontrado en los datos de la fila {index + 1}. Revise los nombres de las columnas en Excel y la plantilla ZPL.")
        except Exception as e:
            print(f"Ocurrió un error al procesar la entrada #{index + 1}: {e}")

    print("\n--- Proceso de impresión desde Excel completado ---")

if __name__ == "__main__":
    # Configuración de la impresora
    ip_impresora = "172.16.1.203"  # IP de la impresora
    puerto_impresora = 9100       # Puerto Telnet 
    
    RUTA_EXCEL = r"C:\proyectos\etiquetas.xlsx"
            
    codigo_a_enviar = """ 
                ^XA
                ^CI28
                ^PW560
                ^LL2435
                ^POI
                ^LH20,1440

                ^FO20,0
                ^A0N,45,45
                ^FDPATATAS HIJOLUSA S.L.^FS

                ^FO20,80
                ^A0N,35,35
                ^FD{DescProducto}^FS

                ^FO20,150
                ^A0N,30,30
                ^FDProveedor:^FS
                ^FO200,150
                ^A0N,30,30
                ^FD{DescProveedor}^FS

                ^FO20,210
                ^A0N,30,30
                ^FDLote:^FS
                ^FO200,210
                ^A0N,30,30
                ^FD{Lote}^FS

                ^FO20,270
                ^A0N,30,30
                ^FDPaquete:^FS
                ^FO200,270
                ^A0N,30,30
                ^FD{Paquete}^FS

                ^FO20,330
                ^A0N,30,30
                ^FDFecha:^FS
                ^FO200,330
                ^A0N,30,30
                ^FD{Fecha}^FS

                ^FO20,390
                ^A0N,30,30
                ^FDVariante:^FS
                ^FO200,390
                ^A0N,30,30
                ^FD{Variante}^FS

                ^FO20,450
                ^A0N,30,30
                ^FDAgricultor:^FS
                ^FO200,450
                ^A0N,30,30
                ^FD{Agricultor}^FS

                ^FO20,510
                ^A0N,30,30
                ^FDOrigen:^FS
                ^FO200,510
                ^A0N,30,30
                ^FD{Origen}^FS

                ^FO40,580
                ^BY2,2,100
                ^BCN,100,Y,N,N
                ^FD{Lote}-{Paquete}^FS

                ^XZ
                """
    
    asyncio.run(imprimir_desde_excel(RUTA_EXCEL, ip_impresora, puerto_impresora, codigo_a_enviar))
