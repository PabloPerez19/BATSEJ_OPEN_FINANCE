import sqlite3
import pandas as pd
import os
import sys
import subprocess
import win32com.client as win32


# Función para enviar correo
def enviar_correo(ruta_archivo, destinatario):
    """
    Envía un correo electrónico con el archivo adjunto usando Outlook.

    Parámetros:
    ruta_archivo (str): Ruta del archivo a adjuntar.
    destinatario (str): Correo electrónico del destinatario.
    """
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = "Resultados de Comisiones - BATSEJ OPEN FINANCE"
        mail.Body = "Adjunto los resultados de comisiones para julio y agosto de 2024."
        mail.Attachments.Add(ruta_archivo)
        mail.Send()
        print(f"✅ Correo enviado a: {destinatario}")
    except Exception as e:
        print(f"❌ Error al enviar el correo: {e}")

# Función para calcular comisiones
def calcular_comision(empresa, peticiones_exitosas, peticiones_no_exitosas):
    """
    Calcula la comision, el iva y el valor total segun la empresa y el número de peticiones.

    Parámetros:
    empresa (str): Nombre de la empresa.
    peticiones_exitosas (int): Número de peticiones exitosas.
    peticiones_no_exitosas (int): Número de peticiones no exitosas.

    Retorna:
    tupla con (comisión, IVA, valor total)
    """
    comision = 0

    # Calculamos la comisión según el contrato de cada empresa
    if empresa == "Innovexa Solutions":
        comision = peticiones_exitosas * 300
    elif empresa == "NexaTech Industries":
        if peticiones_exitosas <= 10000:
            comision = peticiones_exitosas * 250
        elif 10001 <= peticiones_exitosas <= 20000:
            comision = (10000 * 250) + ((peticiones_exitosas - 10000) * 200)
        else:
            comision = (10000 * 250) + (10000 * 200) + ((peticiones_exitosas - 20000) * 170)
    elif empresa == "QuantumLeap Inc.":
        comision = peticiones_exitosas * 600
    elif empresa == "Zenith Corp.":
        if peticiones_exitosas <= 22000:
            comision = peticiones_exitosas * 250
        else:
            comision = (22000 * 250) + ((peticiones_exitosas - 22000) * 130)
    elif empresa == "FusionWave Enterprises":
        comision = peticiones_exitosas * 300

    # Aplicamos descuentos sobre la comision antes del iva
    if empresa == "Zenith Corp." and peticiones_no_exitosas > 6000:
        descuento = comision * 0.05
        comision -= descuento
    elif empresa == "FusionWave Enterprises":
        if 2500 <= peticiones_no_exitosas <= 4500:
            descuento = comision * 0.05
            comision -= descuento
        elif peticiones_no_exitosas > 4500:
            descuento = comision * 0.08
            comision -= descuento

    # Calculamos el iva 19%
    iva = comision * 0.19

    # Calculamos el valor total (comisión + iva)
    valor_total = comision + iva

    return comision, iva, valor_total

# Función para procesar un mes
def procesar_mes(df, mes):
    """
    Procesa los datos de un mes específico para cada empresa.

    Parámetros:
    df (DataFrame): Datos filtrados por empresa y peticiones.
    mes (int): Mes a procesar (7 para julio, 8 para agosto).

    Retorna:
    DataFrame: Resultados con las comisiones calculadas.
    """
    resultados_mes = []
    
    # Filtrar por mes
    df_mes = df[df['date_api_call'].dt.month == mes]
    
    # Calcular comisiones para cada empresa en el mes
    for empresa, grupo in df_mes.groupby('commerce_name'):
        peticiones_exitosas = len(grupo[grupo['ask_status'] == 'Successful'])
        peticiones_no_exitosas = len(grupo[grupo['ask_status'] == 'Unsuccessful'])
        
        comision, iva, valor_total = calcular_comision(empresa, peticiones_exitosas, peticiones_no_exitosas)
        
        # Buscamos el NIT y correo de la empresa
        commerce_nit = grupo['commerce_nit'].iloc[0]  # Tomamos el primer valor
        commerce_email = grupo['commerce_email'].iloc[0] 
        
        resultados_mes.append({
            'Fecha-Mes': mes,
            'Nombre': empresa,
            'Nit': commerce_nit,
            'Valor_comision': comision,
            'valor_iva': iva,
            'Valor_total': valor_total,
            'Correo': commerce_email
        })
    
    return pd.DataFrame(resultados_mes)

# Ejecutar todo
if __name__ == "__main__":
    
    # Conectar a la base de datos
    conn = sqlite3.connect('database.sqlite')
    df_apicall = pd.read_sql_query("SELECT * FROM apicall;", conn)
    df_commerce = pd.read_sql_query("SELECT * FROM commerce;", conn)
    
    # Combinar y filtrar datos
    df = pd.merge(df_apicall, df_commerce, on='commerce_id')
    df = df[df['commerce_status'] == 'Active']
    df['date_api_call'] = pd.to_datetime(df['date_api_call'])
    df = df[df['date_api_call'].dt.year == 2024]

    # Procesar julio y agosto
    julio = procesar_mes(df, 7)
    agosto = procesar_mes(df, 8)
    df_resultados = pd.concat([julio, agosto], ignore_index=True)

    # Crear carpeta de resultados si no existe
    if not os.path.exists('resultado'):
        os.makedirs('resultado')

    # Guardar resultados en Excel
    ruta_archivo = os.path.join('resultado', 'resultados_comisiones.xlsx')
    df_resultados.to_excel(ruta_archivo, index=False)
    print(f"\n✅ Resultados guardados en: {ruta_archivo}")

    # Enviar correo con los resultados
    enviar = input("\n¿Quieres enviar los resultados por correo? (sí/no): ").strip().lower()
    if enviar == "sí" or enviar == "si":
        destinatario = input("Ingresa el correo del destinatario: ").strip()
        enviar_correo(ruta_archivo, destinatario)
