# marca-fantasy-xlsx

Programa para descargar información y estadísticas de jugadores del juego Fantasy Marca. La descarga se realiza en uno o varios ficheros Excel dependiendo de los parámetros especificados por línea de comandos.

## Instalación

El programa está escrito en Ruby y por lo tanto necesita su intérprete para poder se ejecutado. En linux se puede instalar desde los repositorios de cada distribución. En otros sistemas puedes usar [rvm](https://rvm.io/rvm/install), o puedes descargarlo desde la web del propio lenguaje [Ruby](https://www.ruby-lang.org/en/downloads/).

Es necesario instalar la gem [write_xlsx](https://github.com/cxn03651/write_xlsx/tree/master) para la generación del fichero excel. Puedes hacerlo con el comando:`

`$ gem install write_xlsx`

También es necesario instalar la gem [down](https://github.com/janko/down) para la descarga de imágenes desde las API de La Liga. Puedes hacerlo con el comando:

`$ gem install down`

También es necesatio instalar la gem [optparse](https://github.com/skeeto/optparse) para la gestión de los parámetros de línea de comandos. Puedes hacerlo con el siguiente comando:

`$ gem install optparse`

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
    
  - **-m, --compare_players**
 
    Comparación de jugadores. Se usa con la opción **--players_file** para generar un fichero que contiene información estadística de jugadores para su comparación. 
    Es recomendable usarla con **-i, --players_id LIST** y/o **-s, --search_names LIST** (ambos filtros pueden funcionar conjuntamente) para restringir la comparación
    a pocos jugadores.
    
 - **-l, --simulate_team_file FILE**
 
    Simulación de alineaciones. Tiene como parámetro el fichero excel que contendrá la simulación de las alineaciones. Realiza una simulación de alineaciones y calcula     la alineación óptima según varios criterios de puntuación: Puntuación media, puntuación media en unas jornadas determinadas (necesita la opción --weeks),               Puntuación máxima y puntuación de la última jornada. Genera también una hoja dinámica donde simular manualmente las alineaciones calculando las puntuaciones           obtenidas por el equipo. Solo tiene en cuenta los jugadores dispobibles, los lesionados, sancionados y dudosos son excluidos de la simulación.
    
 - **-l, --teams_file FILE**
 
    Fichero de equipos. Tiene como parámetro un fichero JSON que contiene los jugadores del equipo. En la carpeta de ejemplos se puede consultar la estructura del         fichero.
    
 - **-q, --include_questionable_players**
 
    Incluye los jugadores dudosos en la simulación de alineaciones. 
    
  - **-w, --weeks LISTA**
 
    Filtro de jornadas. Permite especificar qué jornadas se mostrarán en los ficheros excel. Se usa con la función de comparación (**--players_file** y 
    **--compare_players** conjuntamente) o con **--players_folder FOLDER** ya que son las opciones que descargan información estadística por jornadas.
    
  - **-c, --chart**
 
    Generación de gráficos. Con esta opción se añadirá al fichero excel una pestaña que contiene los gráficos relacionados con las estadísticas. Se usa con la 
    función de comparación (**--players_file** y **--compare_players** conjuntamente) o con **--players_folder** ya que son las opciones que descargan              
    información estadística por jornadas.
  
## Ejemplos

`$ ruby marca-fantasy-xlsx.rb --players_file /usr/local/data/fantasy.xlsx`

Descarga la lista de jugadores en el fichero /usr/local/data/fantasy.xlsx

`$ ruby marca-fantasy-xlsx.rb --players_folder /usr/local/data/fantasy`

Descarga la información estadística de todos los jugadores en la carpeta /usr/local/data/fantasy. Genera un fichero excel por cada jugador.

`$ ruby marca-fantasy-xlsx.rb --players_folder /usr/local/data/fantasy --players_id 68,99,200`

Descarga la información estadística de los jugadores cuyos identificadores son 68, 99 y 200 en la carpeta /usr/local/data/fantasy. Genera un fichero excel por cada jugador.

`$ ruby marca-fantasy-xlsx.rb --players_folder /usr/local/data/fantasy --players_names isi,galarreta,ledesma`

Descarga la información estadística de los jugadores cuyos nombres contienen isi, galarreta y ledesma en la carpeta /usr/local/data/fantasy. Genera un fichero excel por cada jugador.

`$ ruby marca-fantasy-xlsx.rb --players_file /usr/local/data/fantasy.xlsx --players_names isi,galarreta,ledesma --compare_players --weeks 5..10 --chart`

Descarga la información estadística de los jugadores cuyos nombres contienen isi, galarreta y ledesma en el fichero /usr/local/data/fantasy. Genera un fichero excel que contiene información estadística de las jornadas 5 a la 10 de todos los jugadores para su comparacion y una pestaña de gráficos de los datos estadísticos.

`$ ruby marca-fantasy-xlsx.rb --simulate_team_file /usr/local/data/fantasy.xlsx --teams_file teams.json --weeks 7..13`

Realiza una simulación de alineaciones y guarda el resultado en el fichero /usr/local/data/fantasy.xlsx usando los jugadores del fichero teams.json y calculando las medias de puntos de las jornadas 7 a 13.
