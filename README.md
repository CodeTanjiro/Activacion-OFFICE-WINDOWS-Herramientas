1. **Act. Método 1 (Licencia en Volumen) - Microsoft 365**: Este script activa Microsoft Office utilizando una licencia en volumen a través de un método específico. Descarga una licencia en formato .xrm-ms desde la carpeta de licencias del sistema y la instala. Luego, configura el servidor de activación y realiza la activación del producto.

2. **Act. Método 2 (Licencia en Volumen) - Externo**: Activa Microsoft Office utilizando una licencia en volumen obtenida desde un servidor externo. Descarga la licencia desde una URL remota y la instala en el sistema, lo que facilita la activación del producto sin depender de los servidores de Microsoft.

3. **Act. Método 3 (Reactivación) - Microsoft 365**: Reactiva Microsoft Office en el caso de que la activación previa haya caducado o se necesite volver a activar por algún motivo. Este script ejecuta el proceso de activación nuevamente para restablecer el estado activado del producto.

4. **Activación de Windows 10 (Licencia en Volumen) - Microsoft**: Permite la activación de Windows 10 utilizando una clave de producto específica para cada edición del sistema operativo. Después de que el usuario elige la edición correspondiente, el script inserta la clave de producto adecuada y procede con el proceso de activación.

5. **Eliminar clave de producto de Office**: Este script elimina la clave de producto de Microsoft Office actualmente instalada en el sistema. Proporciona una forma de limpiar las claves de producto previas antes de ingresar una nueva.

6. **Verificar activación de Windows y Office**: Muestra el estado de activación tanto de Windows como de Office. Utiliza el script "ospp.vbs" para verificar la activación de Office y el comando "slmgr /xpr" para verificar la activación de Windows. Proporciona una rápida comprobación del estado de activación de ambos productos.

7. **Convertir archivos a una extensión específica**: Este script toma archivos de una carpeta especificada por el usuario y los convierte a una extensión específica también especificada por el usuario. Los archivos se copian con la nueva extensión en la misma carpeta.

8. **Verificar puertos abiertos**: Escanea los puertos en estado "ESTABLISHED" y guarda la información sobre la dirección remota y el puerto remoto en un archivo llamado "hosts_abiertos.txt". Proporciona una forma de monitorear los puertos abiertos en el sistema.

9. **Ver mi dirección IP**: Muestra la dirección IP detallada del sistema y la guarda en un archivo llamado "direccion_ip.txt". Brinda una forma rápida de obtener y guardar la información sobre la dirección IP del sistema.

10. **Guardar especificaciones de la PC**: Recopila información detallada sobre la configuración y especificaciones de la computadora utilizando el comando "systeminfo" y guarda los resultados en un archivo llamado "Especificaciones_PC.txt". Ofrece una manera conveniente de obtener un informe completo de las especificaciones de hardware y software del sistema.

Las credenciales de acceso al sistema son "admin" tanto para el nombre de usuario como para la contraseña. Estas credenciales son necesarias para acceder al menú principal y utilizar cualquiera de las opciones proporcionadas en el script.
