import sqlite3
import pandas as pd
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    
)

def connect_to_db(db_path):
    """Conecta a la base de datos SQLite."""
    try:
        conn = sqlite3.connect(db_path)
        logging.info(f"Conectado exitosamente a la base de datos: {db_path}")
        return conn
    except sqlite3.Error as e:
        logging.error(f"Error al conectar a la base de datos: {e}")
        raise

def load_data(conn):
    """Carga los datos necesarios de la base de datos."""
    query = """
    WITH monthly_calls AS (
        SELECT 
            c.commerce_name, 
            c.commerce_nit, 
            c.commerce_email, 
            strftime('%Y-%m', a.date_api_call) as month,
            SUM(CASE WHEN a.ask_status = 'Successful' THEN 1 ELSE 0 END) as successful_calls,
            SUM(CASE WHEN a.ask_status = 'Unsuccessful' THEN 1 ELSE 0 END) as unsuccessful_calls
        FROM commerce c
        JOIN apicall a ON c.commerce_id = a.commerce_id
        WHERE c.commerce_status = 'Active' AND 
              a.date_api_call BETWEEN '2024-07-01' AND '2024-08-31'
        GROUP BY c.commerce_name, c.commerce_nit, c.commerce_email, month
    )
    SELECT 
        month as 'Fecha-Mes',
        commerce_name as 'Nombre',
        commerce_nit as 'Nit',
        successful_calls,
        unsuccessful_calls,
        commerce_email as 'Correo',
        CASE
            WHEN commerce_name = 'Innovexa Solutions' THEN successful_calls * 300
            WHEN commerce_name = 'NexaTech Industries' THEN
                CASE
                    WHEN successful_calls <= 10000 THEN successful_calls * 250
                    WHEN successful_calls <= 20000 THEN successful_calls * 200
                    ELSE successful_calls * 170
                END
            WHEN commerce_name = 'QuantumLeap Inc.' THEN successful_calls * 600
            WHEN commerce_name = 'Zenith Corp.' THEN
                CASE
                    WHEN successful_calls <= 22000 THEN successful_calls * 250
                    ELSE successful_calls * 130
                END
            WHEN commerce_name = 'FusionWave Enterprises' THEN successful_calls * 300
            ELSE 0
        END as 'Valor_comision_base',
        CASE
            WHEN commerce_name = 'Zenith Corp.' AND unsuccessful_calls > 6000 THEN 0.05
            WHEN commerce_name = 'FusionWave Enterprises' AND unsuccessful_calls BETWEEN 2500 AND 4500 THEN 0.05
            WHEN commerce_name = 'FusionWave Enterprises' AND unsuccessful_calls > 4501 THEN 0.08
            ELSE 0
        END as 'Porcentaje_descuento'
    FROM monthly_calls
    """

    try:
        df = pd.read_sql_query(query, conn)
        logging.info(f"Datos cargados exitosamente. Forma del DataFrame: {df.shape}")
        return df
    except pd.io.sql.DatabaseError as e:
        logging.error(f"Error al ejecutar la consulta SQL: {e}")
        raise

def process_data(df):
    """Procesa los datos y calcula las comisiones finales."""
    logging.info("Procesando datos y calculando comisiones finales")

    df["Descuento"] = df["Valor_comision_base"] * df["Porcentaje_descuento"]
    df["Valor_comision"] = df["Valor_comision_base"] - df["Descuento"]
    df["Valor_iva"] = df["Valor_comision"] * 0.19
    df["Valor_Total"] = df["Valor_comision"] + df["Valor_iva"]

    columns_order = [
        "Fecha-Mes",
        "Nombre",
        "Nit",
        "Valor_comision",
        "Descuento",
        "Valor_iva",
        "Valor_Total",
        "Correo",
    ]
    df_final = df[columns_order]

    logging.info(f"Datos procesados. Forma del DataFrame final: {df_final.shape}")
    return df_final

def export_to_excel(df, filename):
    """Exporta los resultados a un archivo Excel."""
    try:
        df.to_excel(filename, index=False, engine="openpyxl")
        logging.info(f"Resultados exportados exitosamente a: {filename}")
    except Exception as e:
        logging.error(f"Error al exportar a Excel: {e}")
        raise

def send_email(recipient, subject, body, attachment):
    """Envía un correo electrónico con los resultados."""
    sender_email = "tu_correo@ejemplo.com"
    password = "tu_contraseña"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    with open(attachment, "rb") as attachment_file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment_file.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(attachment)}",
    )
    message.attach(part)

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(message)
        logging.info(f"Correo enviado exitosamente a {recipient}")
    except Exception as e:
        logging.error(f"Error al enviar el correo: {e}")
        raise


def main():
    """Función principal que ejecuta todo el proceso."""
    db_path = r"C:\Users\alejmora\Documents\prueba_banco\database.sqlite"
    excel_output = "resultados_comisiones_final.xlsx"

    try:
        conn = connect_to_db(db_path)
        raw_data = load_data(conn)
        if raw_data.empty:
            logging.warning("No se encontraron datos en la base de datos.")
            return

        processed_data = process_data(raw_data)
        export_to_excel(processed_data, excel_output)

        recipient_email = (
            "alejomdg@gmail.com" 
        )
        subject = "Resultados de comisiones"
        body = "Adjunto encontrará los resultados de las comisiones."
        send_email(recipient_email, subject, body, excel_output)
        
        logging.info("Proceso completado exitosamente.")

        print("\nResumen de resultados:")
        print(processed_data.describe())
        print("\nPrimeras filas del DataFrame final:")
        print(processed_data.head())

    except Exception as e:
        logging.error(f"Error en la ejecución del proceso: {e}")
    finally:
        if conn:
            conn.close()
            logging.info("Conexión a la base de datos cerrada")

if __name__ == "__main__":
    main()
