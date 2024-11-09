import os
from flask import Flask, send_file
import paramiko
from io import BytesIO

# Crear una instancia de Flask
app = Flask(__name__)

# Parámetros de conexión SSH
hostname = 'ssh-natureza.alwaysdata.net'  # Dirección del servidor SSH
port = 22  # Puerto SSH
username = 'natureza_anon'  # Tu nombre de usuario en el servidor
password = '(123456)'  # Tu contraseña SSH

# Ruta remota donde está el archivo Excel en el servidor
ruta_remota = 'PEREZ.xlsx'


# Función para obtener el archivo Excel desde el servidor remoto
def obtener_archivo_excel_remoto():
    # Crear una instancia SSHClient
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(
        paramiko.AutoAddPolicy())  # Agregar automáticamente la clave del host

    # Conectarse al servidor remoto
    ssh.connect(hostname, port=port, username=username, password=password)

    # Usar SFTP para leer el archivo desde el servidor remoto
    sftp = ssh.open_sftp()

    # Leer el archivo remoto en memoria
    with sftp.open(ruta_remota, 'rb') as remote_file:
        archivo_excel = remote_file.read()

    # Cerrar la conexión SFTP y SSH
    sftp.close()
    ssh.close()

    # Devolver el contenido del archivo Excel (en memoria)
    return archivo_excel


@app.route('/')
def index():
    return "Para descargar el archivo, ve a <a href='/download'>/download</a>."


@app.route('/download')
def download():
    # Obtener el archivo Excel desde el servidor remoto
    archivo_excel = obtener_archivo_excel_remoto()

    # Crear un objeto BytesIO a partir del contenido del archivo Excel
    excel_buffer = BytesIO(archivo_excel)

    # Servir el archivo para su descarga
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name="PEREZ.xlsx",
        mimetype=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    # Ejecutar la aplicación Flask
    app.run(debug=True, host="0.0.0.0", port=5000)
