import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import os
from config import EMAIL_CONFIG

# Códigos de color ANSI
VERDE = '\033[92m'
AZUL = '\033[94m'
RESET = '\033[0m'

def enviar_correo():
    # Configuración del correo desde el archivo de configuración
    remitente = EMAIL_CONFIG['remitente']
    destinatarios = EMAIL_CONFIG['destinatarios']
    contraseña = EMAIL_CONFIG['contraseña']
    ruta_excel = EMAIL_CONFIG['ruta_excel']
    
    # Crear el mensaje
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = ', '.join(destinatarios)  # Unir todos los destinatarios con comas
    
    # Obtener fecha y hora actual
    fecha_hora = datetime.now()
    asunto = f"Resumen Restó | {fecha_hora.strftime('%d/%m/%Y')} | {fecha_hora.strftime('%H:%M')}"
    msg['Subject'] = asunto
    
    # Cuerpo del correo
    cuerpo = f"""
    Buen día,
    
    Adjunto encontrará el resumen del día {fecha_hora.strftime('%d/%m/%Y')}.
    
    Saludos cordiales,
    Rio Arriba Restó
    """
    
    msg.attach(MIMEText(cuerpo, 'plain'))
    
    # Adjuntar el archivo Excel
    try:
        if os.path.exists(ruta_excel):
            # Determinar el tipo MIME basado en la extensión del archivo
            extension = os.path.splitext(ruta_excel)[1].lower()
            if extension == '.xlsx':
                mime_type = 'xlsx'
            elif extension == '.ods':
                mime_type = 'ods'
            else:
                print(f"Error: Formato de archivo no soportado: {extension}")
                return
                
            with open(ruta_excel, 'rb') as f:
                excel_adjunto = MIMEApplication(f.read(), _subtype=mime_type)
                excel_adjunto.add_header('Content-Disposition', 'attachment', 
                                       filename=os.path.basename(ruta_excel))
                msg.attach(excel_adjunto)
        else:
            print(f"Error: No se encontró el archivo en {ruta_excel}")
            print("Por favor, verifica que:")
            print("1. El archivo existe en la ruta especificada")
            print("2. El nombre del archivo es correcto")
            print("3. Tienes permisos para acceder al archivo")
            return
    except Exception as e:
        print(f"Error al adjuntar el archivo: {str(e)}")
        return
    
    # Enviar el correo
    try:
        # Conectar al servidor SMTP de Gmail
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(remitente, contraseña)
        
        # Enviar el correo a todos los destinatarios
        servidor.send_message(msg)
        servidor.quit()
        
        print("\n" + "="*50)
        print("Correo enviado exitosamente!")
        print(f"\n{VERDE}INFORMATICA ORIGEN")
        print("LIC. FRANCO MONZON")
        print(f"{AZUL}Contacto: 3795-166911")
        print(f"{RESET}" + "="*50 + "\n")
        
    except Exception as e:
        print(f"Error al enviar el correo: {str(e)}")

if __name__ == "__main__":
    enviar_correo() 