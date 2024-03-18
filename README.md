## Imperium Clasico

Clon de ImperiumAO ambientado en las versiones clásicas (1.3 & 1.4) [Mod AOLibre]

- Se paso el control de mucho de los sistemas que no debian de estar en el cliente al servidor.
- Utiliza 2 servidores SQL, uno para cuentas y otro para personajes.
- Se diseño y programo un nuevo sistema de cuentas que se controla totalmente desde el servidor, lo que mejora muchisimo la seguridad.
- Se utiliza el Batch para la mejora del motor grafico, se mejoraron varios aspectos del codigo de Argentum, entre otras caracteristicas.
- Sistema de familiares
- Sistema de Creditos
- Todas las interfaces adaptadas.
- Sistema de macros (se guardan en el PJ)
- Macro de trabajo controlado desde servidor.
- Sistema de casamientos
- Particulas, luces, etc...
- Anillos y otros items magicos.
Y otras caracterisiticas.

** Liberación por descontinuación del proyecto **

NOTA: En este proyecto se estaba intentando fusionar mapas de la versión 1.3 y 1.4, pero no se llego a terminar.

Repositorio de la web: https://github.com/Lorwik/ImperiumClasico-WebAngular

# Instrucciones:

Para abrir el servidor es necesario contar con una base de datos MySQL y MySQL ODBC 8.0 ANSI Driver.
En la carpeta Configuracion del servidor encontraras un archivo llamado DataBase.ini, ahi deberas configurar los datos para la conexión a la base de datos. En la carpeta Fixtures encontraras 2 archivos SQL que corresponde a las cuentas/web y a los personajes.

Para configurar la IP y el puerto debes crear un archivo .txt con el siguiente contenido y subirlo a la web:

127.0.0.1|7666|LocalHost;127.0.0.1|7666|Servidor Secundario

IP|PUERTO|NOMBRE DEL SERVIDOR;

En el Mod_General del cliente busca:

Public Sub ListarServidores()

y en la siguiente linea cambia la URL donde alojas el archivo .txt:

responseServer = Inet.OpenRequest("https://tuurl.com/server-listiac.txt", "GET")

# Creditos:

By Lorwik
