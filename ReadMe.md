Librerias utilizadas:

	- os: Para funciones del sistema operativo, como lectura y escritura de archivos.
	- glob: Se utiliza para buscar patrones entre los nombres de los archivos.
	- pandas: Para la manipulacion e interacción con documentos .xls y .xlsx
	- openpyxl: Para la manipulacion e interaccion con documentos .xlsx
	- mariadb: Para la base de datos del script.

Consideraciones para utilizar el script:

	- Para instalar las diferentes dependencias, si se utiliza Linux:

		- pip3 install pandas
		- pip3 install openpyxl
		- pip3 install mariadb

	Consideraciones para la utilización de la base de datos:

		- Por defaut vienen las credenciales de origen, es decir, las de mi equipo, es necesario cambiar las variables
		globales, para obtener un correcto funcionamiento. Se encuentran al inicio del archivo.

Para correr el script:
	
	$ python3 Excel.py

Comentarios:
Me pareció un ejercicio lo bastante robusto para explotar distintos aspectos, pero a su vez, no tan complicado,
si bien me atore un poco, espero que todo este correcto, y sin mas, espero que todo se encuentre en orden.