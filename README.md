# marca-fantasy-xlsx

Programa para descargar información y estadísticas de jugadores del juego Fantasy Marca. La descarga se realiza en uno o varios ficheros Excel dependiendo de los parámetros especificados por línea de comandos.

## Instalación

El programa está escrito en Ruby y por lo tanto necesita su intérprete para poder se ejecutado. En linux se puede instalar desde los repositorios de cada distribución, puedes usar [rvm](https://rvm.io/rvm/install), o puedes descargarlo desde la web del propio lenguaje [Ruby](https://www.ruby-lang.org/en/downloads/).

Es necesario instalar la gem [write_xlsx](https://github.com/cxn03651/write_xlsx/tree/master) para la generación del fichero excel. Puedes hacerlo con el comando:`

`$ gem install write_xlsx`

También es necesario instalar la gem [down](https://github.com/janko/down) para la descarga de imágenes desde las API de La Liga. Puedes hacerlo con el comando:

`$ gem install down`

## Ejecución

Existen varias opciones de ejecución del programa que se pueden listar con el comando:

`$ ruby marca-fantasy-xlsx.rb --help`

Las opciones de ejecución son las siguientes:

- **-f, --players_file FILE**

  Especifica el nombre del fichero donde se generará el fichero de jugadores. Sin más opciones creará un fichero excel que contiene una tabla dinámica con todos los       jugadores
- **-p, --players_folder FOLDER**

  Especifica el nombre de la carpeta donde se generarán los ficheros de información estadística de los jugadores. Sin más opciones creará un fichero excel por cada       jugador.
  
 - **-i, --players_id LIST**
 
  Filtrado de jugadores por su identificador. Se usa para especificar qué jugadores se descargarán desde la API de Fantasy Marca. Se puede utilizar con las opciones
  **--players_folder** y **--players_file**. El parámetro consiste en una lista de identificadores separados por comas y sin espacios. Los idenfificadores se pueden
  consultar en el fichero generado con la opción **--players_file**.
  
 - **-s, --search_names LIST**
 
  Filtrado de jugadores por su nombre. Se usa para especificar qué jugadores se descargarán desde la API de Fantasy Marca. Se puede utilizar con las opciones
  **--players_folder** y **--players_file**. El parámetro consiste en una lista de identificadores separados por comas y sin espacios. No es necesario indicar el
  nombre completo del jugador, indicando una parte del nombre ya realiza la descarga. 
