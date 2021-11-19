require './fantasy_constants.rb'
require './fantasy.rb'
require './player.rb'
require 'i18n'

#
# Clase que contiene una lista de jugadores
#

class Players < Fantasy

   #
   # Constructor. Descarga la lista de jugadores basándose en las opciones especificadas en la línea de comandos.
   #
   # @raise [Timeout::Error, Errno::EINVAL, Errno::ECONNRESET, EOFError,
   #          Net::HTTPBadResponse, Net::HTTPHeaderSyntaxError, Net::ProtocolError] Error en la descarga del fichero JSON
   # @author Gerard Carrasquer
   #

   def initialize
      I18n.config.available_locales = :en
      # Obtención de la lista de jugadores
      @players_data = get_data("#{FANTASY_API_SERVER}/#{FANTASY_PLAYERS_URL}")
      @players = Array.new
      
      super()
      
      if @players_data[JSON_RESPONSE] == JSON_ERROR
         # Error en la obtención del fichero
         puts "Error: #{@players_data[JSON_DATA][JSON_MESSAGE]}"
      else
         if !$options[:teams_file].nil?
            begin
               f_file = File.open($options[:teams_file])
               @teams_data = JSON.load(f_file)
               f_file.close
            rescue => e
               puts e.to_s
               if $options[:verbose]
                  puts "#{e.to_s}\n\n#{e.backtrace}"
               end
            end
         end
            
         if (!$options[:players_file].nil? && $options[:compare_players]) || !$options[:players_folder].nil? || !$options[:simulate_team_file].nil?
            # Opción de descarga de la lista de jugadores o de descarga de fichas de jugadores
            # Ambas opciones admiten el filtrado de jugadores
            @players_data[JSON_DATA].each do |p|
               if filter(p[JSON_ID], p[JSON_NICKNAME])
                  #El jugador cumple los criterios especificados por el usuario en línea de comandos
                  print "Descargando jugador id #{p[JSON_ID]}: #{p[JSON_NICKNAME]}.... "
                  player = Player.new(p[JSON_ID])
                  print "#{player[JSON_RESPONSE]}\n"
                  
                  @players.push(player)
               end
            end
         end
         
         if !$options[:players_file].nil?
            # Se ha especificado un fichero para descarga
            if $options[:compare_players]
               # Se ha especificado la comparación de jugadores
               to_compare_xlsx($options[:players_file])
            else
               # Se ha especificado la descarga de la lista jugadores
               to_xlsx($options[:players_file])
            end
         end
         if !$options[:players_folder].nil?
            # Se ha especificado la descarga de fichas de jugadores en una carpeta
            @players.each do |p|
               # Se crea un fichero para cada jugador
               p.to_xlsx("#{$options[:players_folder]}/#{I18n.transliterate(p[JSON_DATA][JSON_NICKNAME]).gsub(/[[:blank:]]/, '_')}.xlsx")
            end
         end
         
         if !$options[:simulate_team_file].nil?
            # Se ha especificado un fichero para simulación
            to_simulate_xlsx($options[:simulate_team_file])
         end
         
      end
   end
   
   #
   # Creación del fichero excel que contiene la lista de jugadores.
   #
   # @param p_file_name [String] Nombre del fichero 
   #
   # @raise [Exception] Error en la escritura del fichero excel
   # @author Gerard Carrasquer
   #
   
   def to_xlsx(p_file_name)
      
      #Text formats
      f_text_color = '#2e4053'
      f_text_bg_color = '#fef9e7'
      f_info_color = '#34495e'
      f_header_color = '#34495e'
      f_header_bg_color = '#eafaf1'
      f_info_bg_color = '#ebf5fb'
      f_text_font = 'Calibri'
      
      print "Generando fichero #{p_file_name}... "
            
      f_workbook = WriteXLSX.new(p_file_name)
      
      # Formatos
      f_name_format = f_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
      f_info_format = f_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_info_bg_color, :border => 1)
      f_header_format = f_workbook.add_format(:color => f_header_color, :size => 10, :font => f_text_font, 
                        :bg_color => f_header_bg_color, :border => 1, :bold => 1, :align => 'justify')
      f_currency_format = f_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font, 
                        :num_format => '#,##0', :bg_color => f_info_bg_color, :border => 1)
      

      f_worksheet = f_workbook.add_worksheet('Jugadores')
      
      f_worksheet.merge_range('B2:I2', 'Jugadores', f_name_format)
     
      # Establecemos el ancho de las columnas
      
      f_worksheet.set_column('B:B', 5) 
      f_worksheet.set_column('C:D', 15)
      f_worksheet.set_column('E:E', 10)
      f_worksheet.set_column('F:G', 13)
      f_worksheet.set_column('H:I', 10)
      
      # Establecemos el alto de las filas
      
      f_worksheet.set_row(4, 35)
      f_worksheet.set_row(1, 20)
    
      f_worksheet.add_table("B5:I#{7 + @players_data[JSON_DATA].length}", :columns => [
                     { :header => 'Id', :header_format => f_header_format },
                     { :header => 'Nombre', :header_format => f_header_format },
                     { :header => 'Equipo', :header_format => f_header_format },
                     { :header => 'Posición', :header_format => f_header_format },
                     { :header => 'Puntos en la última temporada', :header_format => f_header_format },
                     { :header => 'Puntos en la temporada actual', :header_format => f_header_format },
                     { :header => 'Media de puntos', :header_format => f_header_format },
                     { :header => 'Valor de mercado', :header_format => f_header_format }
                 ],
                 :style     => 'Table Style Light 11',
                 :name => 'Jugadores')
      
      f_index = 0
      
      @players_data[JSON_DATA].each do |p|
      
         if filter(p[JSON_ID], p[JSON_NICKNAME])
            # El jugador cumple con los criterios especificados por el usuario en lista de comandos
      
            f_worksheet.write_number("B#{f_index + 6}", p[JSON_ID], f_info_format)
            f_worksheet.write("C#{f_index + 6}", p[JSON_NICKNAME], f_info_format)
            f_worksheet.write("D#{f_index + 6}", p[JSON_TEAM][JSON_NAME], f_info_format)
            
            case p[JSON_POSITION_ID]
                  when '1'
                     f_worksheet.write("E#{f_index + 6}", 'Portero', f_info_format)
                  when '2'
                     f_worksheet.write("E#{f_index + 6}", 'Defensa', f_info_format)
                  when '3'
                     f_worksheet.write("E#{f_index + 6}", 'Mediocampista', f_info_format)
                  when '4'
                     f_worksheet.write("E#{f_index + 6}", 'Delantero', f_info_format)
                  else
                     f_worksheet.write("E#{f_index + 6}", 'Desconocida', f_info_format)
            end
               
            if p[JSON_LAST_SEASON_POINTS].nil?
               f_worksheet.write_number("F#{f_index + 6}", 0, f_info_format)
            else
               f_worksheet.write_number("F#{f_index + 6}", p[JSON_LAST_SEASON_POINTS], f_info_format)
            end
            
            f_worksheet.write_number("G#{f_index + 6}", p[JSON_POINTS], f_info_format)
            f_worksheet.write_number("H#{f_index + 6}", p[JSON_AVERAGE_POINTS].round(2), f_info_format)
            f_worksheet.write_number("I#{f_index + 6}", p[JSON_MARKET_VALUE], f_currency_format)
            
            f_index += 1
         end
            
      end
      
      begin
         f_workbook.close
         
         print "ok\n"
      rescue => e
         print "#{e.to_s}\n"
         if $options[:verbose]
            puts "#{e.to_s}\n\n#{e.backtrace}"
         end
      end
   end
   
   #
   # Creación del fichero excel que contiene las estadísticas de varios jugadores por jornadas.
   # Genera un fichero que contiene una tabla para cada estadística existente en el JSON descargado para el jugador.
   # Cada tabla tiene los jugadores en las columnas y las jornadas en las filas. Si se especifica la opción de 
   # generar gráficos en línea de comandos los genera en otra pestaña.
   #
   # @param p_file_name [String] Nombre del fichero 
   #
   # @raise [Exception] Error en la escritura del fichero excel
   # @author Gerard Carrasquer
   #
   
   def to_compare_xlsx(p_file_name)
   
      @stats_columns = Hash.new
      
      f_workbook = WriteXLSX.new(p_file_name)
      
      #Text formats
      f_text_color = '#34495e'
      f_odd_color = '#fcf3cf'
      f_even_color = '#d5f5e3'
      f_text_font = 'Calibri'
      f_odd_format = f_workbook.add_format(:bold => 0, :color => f_text_color, :size => 10,
                  :font => f_text_font, :bg_color => f_odd_color, :border => 1)
      f_even_format = f_workbook.add_format(:bold => 0, :color => f_text_color, :size => 10,
                  :font => f_text_font, :bg_color => f_even_color, :border => 1)
      f_tag_format = f_workbook.add_format(:bold => 1, :color => f_text_color, :size => 10, :font => f_text_font,
                  :border => 1)
      f_title_format = f_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font)
      
      f_small_title_format = f_workbook.add_format(:bold => 1, :color => f_text_color, :size => 12, :font => f_text_font)
   
   
      
      if not @@current_week_number.nil?
         # No ha habido ningún error en la descarga de la jornada actual

         puts "Generando fichero #{p_file_name}... "
         
         f_worksheet = f_workbook.add_worksheet('Estadísticas')
         
         if $options[:chart]
            f_chart_worksheet = f_workbook.add_worksheet('Gráficas')
         end
         
         f_worksheet.set_row(5, 20)
         f_worksheet.set_column('D:Y', 10)
         f_worksheet.set_column('C:C', 15)
         
         f_worksheet.merge_range('B6:F6', 'Estadísticas', f_title_format)
         
         # Columna derecha del borde 
         f_end_border_column = 'B'
         
         # Dos columnas después del último jugador
         1.upto(@players.length + 2) do |i|
            f_end_border_column.succ!
         end
            
                  
         # Estadísticas disponibles en el JSON que hemos descargado y que corresponde al jugador
         f_stat_tags = {JSON_TOTAL_POINTS => "Puntos",
                        JSON_MINS_PLAYED => "Minutos jugados", 
                        JSON_GOALS => "Goles", 
                        JSON_GOAL_ASSIST => "Asistencias de gol", 
                        JSON_GOAL_OFFTARGET_ATT_ASSIST => "Asistencias sin gol", 
                        JSON_PEN_AREA_ENTRIES => "Balones al área", 
                        JSON_PENALTY_WON => "Penaltis provocados", 
                        JSON_PENALTY_SAVE => "Penaltis parados", 
                        JSON_SAVES => "Paradas", 
                        JSON_EFECTIVE_CLEARANCE => "Despejes", 
                        JSON_PENALTY_FAILED => "Penaltis fallados", 
                        JSON_OWN_GOALS => "Goles en propia puerta", 
                        JSON_GOALS_CONCEDED => "Goles en contra", 
                        JSON_YELLOW_CARD => "Tarjetas amarillas", 
                        JSON_SECOND_YELLOW_CARD => "Segunda tarjeta amarilla", 
                        JSON_RED_CARD => "Tarjeta roja", 
                        JSON_TOTAL_SCORING_ATT => "Tiros a puerta", 
                        JSON_WON_CONTEST => "Regates", 
                        JSON_BALL_RECOVERY => "Balones recuperados", 
                        JSON_POSS_LOST_ALL => "Posesiones perdidas", 
                        JSON_PENALTY_CONCEDED =>"Penaltis cometidos",
                        JSON_MARCA_POINTS => "Puntos marca",
                        AVG_DISPERSION => "Desviación de la media"}
         
         f_current_row = 10
            
         f_series = Array.new
         
         f_weeks_count = @@current_week_number
         f_first_week = 1
         if !$options[:weeks].nil? 
            f_weeks_count = $options[:weeks].length
            f_first_week = $options[:weeks][0]
         end
         
         f_stat_tags.each_with_index do |(k,v), i|
         
            # Recorremos todas las estadísticas disponibles para generar un recuadro para cada una que contendrá
            # los jugadores en las columnas y las jornadas en las filas
            
            f_current_column = 'C'
            
            f_worksheet.set_row(f_current_row - 2, 25)
            
            f_chart_row = f_current_row
            
            f_worksheet.merge_range("B#{f_current_row - 1}:F#{f_current_row - 1}", v, f_small_title_format)
            
            draw_spreadsheet_border(f_workbook, f_worksheet, '#3498db', "B#{f_current_row + 1}", 
                        "#{f_end_border_column}#{f_current_row + f_weeks_count + 3}")
            
            f_current_row += 2
            f_worksheet.write_row(f_current_row - 1, 3, @players.map{|p| p[JSON_DATA][JSON_NICKNAME]}, f_tag_format)
            
            f_current_row += 1
            
            f_series.clear
            
            f_chart_type = 'column'
            
            if k == AVG_DISPERSION
               f_chart_type = 'line'
            end
            
            1.upto(@@current_week_number) do |week|
            
               # Recorremos todas las jornadas
            
               if $options[:weeks].nil? || $options[:weeks].include?(week)
                  # Si se ha especificado la opción --weeks en línea de comandos $options[:compare_weeks]
                  # contiene las jornadas a comparar especificadas por el usuario.
                  # En caso contrario $options[:compare_weeks] es nil y debemos recorrer desde la jornada 1
                  # hasta la jornada actual
                  f_current_column = 'C'
                  
                  f_worksheet.write("#{f_current_column}#{f_current_row}", "Jornada #{week}", f_tag_format)
                  
                  @players.each do |p|
                     #@players contiene los jugadores especificados por el usuario en línea de comandos
                  
                     f_current_column.succ!
                     
                     if week == f_first_week
                        # Añadimos la serie de datos para los gráficos una vez por cada jugador
                        # Lo hacemos solo en la primera jornada que recorremos
                        f_series.push({:name => p[JSON_DATA][JSON_NICKNAME], :categories_column => 'C', 
                                 :values_column => f_current_column.clone, 
                                 :init_line => f_current_row, :end_line => f_current_row + f_weeks_count,
                                 :show_value => !(f_chart_type == 'line')})
                     end
                     
                     # Obtenemos las estadísticas del jugador para una jornada determinada
                     f_player_stats = p.get_week_stats(week)
                     
                     
                     if f_player_stats.nil?
                        # La jornada especificada no está en el JSON que hemos descargado, lo cual significa
                        # que el jugador no ha jugado esa jornada. Rellenamos la celda correspondiente a esa jornada
                        # con '-''
                        if k == AVG_DISPERSION
                           f_worksheet.write_number("#{f_current_column}#{f_current_row}", 
                                          -p[JSON_DATA][JSON_AVERAGE_POINTS].round(2),
                                          (f_current_row.odd?) ? f_odd_format : f_even_format)
                        else
                           f_worksheet.write("#{f_current_column}#{f_current_row}", '-',
                                          (f_current_row.odd?) ? f_odd_format : f_even_format)
                        end
                     else
                        # La jornada contiene datos
                        f_data = 0
                        case k
                           when JSON_TOTAL_POINTS
                              # Los puntos totales están en la raiz del subobjeto JSON f_player_stats
                              f_data = f_player_stats[JSON_TOTAL_POINTS]
                           when JSON_MARCA_POINTS
                              # El resto de estadísticas se encuentran dentro del tag JSON_STATS y están representados
                              # como un array de dos posiciones. La posición cero contiene la estadística y la uno
                              # contiene los putos otorgados al jugador para esa estadística
                              f_data = f_player_stats[JSON_STATS][JSON_MARCA_POINTS][1]
                           when AVG_DISPERSION
                              f_data = (f_player_stats[JSON_TOTAL_POINTS] - p[JSON_DATA][JSON_AVERAGE_POINTS]).round(2)  
                           else
                              # k contiene el tag correspondiente a la estadística que estamos tratando
                              f_data = f_player_stats[JSON_STATS][k][0]
                        end
                        
                        f_worksheet.write_number("#{f_current_column}#{f_current_row}", f_data,
                                          (f_current_row.odd?) ? f_odd_format : f_even_format)
                     end
                     
                     
                  end
                  
                  
                  
                  f_current_row += 1
               end
               
            end
            
            if $options[:chart]
               # El usuario ha especificado la opción de generar gráficos
               # Lo insertamos con las series que hemos ido guardando en f_series
               draw_column_chart(f_workbook, f_chart_worksheet, f_worksheet.name, 
                  f_chart_type, "C#{f_chart_row + (i * 22)}", v, 'Jornada', v, *f_series)
                  
            end
            
            f_current_row += 5
         end
      end
      
      begin
         f_workbook.close
      
         puts "Fichero #{p_file_name} generado"
      rescue => e
         print "#{e.to_s}\n"
         if $options[:verbose]
            puts "#{e.to_s}\n\n#{e.backtrace}"
         end
      end
   end
   
   #
   # Establece si un jugador debe ser incluido en la generación de los ficheros basándose en los criterios
   # especificados por el usuario en línea de comandos.
   #
   # @param p_player_id [Integer] Identificador el jugador
   # @param p_player_name [String] Nombre o una parte del nombre el jugador 
   #
   # @return [Bool] El jugador cumple con los criterios de inclusión en los ficheros a generar.
   #
   # @author Gerard Carrasquer
   #
   
   def filter(p_player_id, p_player_name)
      
      # Si ningún filtro está activo la función debe devolver true para tratar todos los jugadores
      f_filter_ok = $options[:players_id].nil? && $options[:players_names].nil? && $options[:teams_file].nil?
       
      if !$options[:players_id].nil?
         f_filter_ok = $options[:players_id].include?(p_player_id)
      end
      
      if !$options[:players_names].nil?
         f_filter_ok |= !$options[:players_names].index{|e| 
                  Regexp.new(".*#{I18n.transliterate(e)}.*", true).match?(I18n.transliterate(p_player_name))}.nil?
      end
      
      if !$options[:teams_file].nil?
         @teams_data.each do |k,v|
            if k == JSON_PLAYERS
               f_filter_ok |= v.include?(p_player_id.to_i) || !v.index{|e| 
               Regexp.new(".*#{I18n.transliterate(e.to_s)}.*", true).match?(I18n.transliterate(p_player_name))}.nil?
            end
         end
      end
      
      return f_filter_ok
   end
   
   #
   # Creación del fichero excel que contiene la simulación de alineaciones.
   #
   # @param p_file_name [String] Nombre del fichero 
   #
   # @raise [Exception] Error en la escritura del fichero excel
   # @author Gerard Carrasquer
   #
   
   def to_simulate_xlsx(p_file_name)
   
      #Text formats
      f_text_color = '#2e4053'
      f_text_font = 'Calibri'
      f_text_bg_color = '#fef9e7'
   
      f_workbook = WriteXLSX.new(p_file_name)
      
      # Formatos
      f_name_format = f_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
            
      
      

      if not @@current_week_number.nil?
         # No ha habido ningún error en la descarga de la jornada actual
         f_goal_keepers = Array.new
         f_defenders = Array.new
         f_midfielders = Array.new
         f_strikers = Array.new
         
         print "Generando fichero #{p_file_name}... "
         
         f_goal_keepers, f_defenders, f_midfielders, f_strikers = get_players_positions
         
         #Generación de la hoja de simulación
         write_simulation_sheet(f_workbook, "Simulación", f_goal_keepers, f_defenders, f_midfielders, f_strikers)
         
         #Generación de la hoja de alineaciones
         write_lineup_sheet(f_workbook, "Alineaciones", f_goal_keepers, f_defenders, f_midfielders, f_strikers)
      
         begin
            f_workbook.close
            
            print "ok\n"
         rescue => e
            print "#{e.to_s}\n"
            if $options[:verbose]
               puts "#{e.to_s}\n\n#{e.backtrace}"
            end
         end
      end
   end
   
   #
   # Distribuye los jugadores en sus posiciones.
   #
   # @return [Array<Hash>] Array que contiene los porteros.
   # @return [Array<Hash>] Array que contiene los defensas.
   # @return [Array<Hash>] Array que contiene los mediocampistas.
   # @return [Array<Hash>] Array que contiene los delanteros.
   # @author Gerard Carrasquer
   #
   
   def get_players_positions
      f_goal_keepers = Array.new
      f_defenders = Array.new
      f_midfielders = Array.new
      f_strikers = Array.new
      
      @players.each do |player|
         
         # Solo tenemos en cuenta los jugadores cuyo estado es ok
         # Si se ha indicado el flag de incluir dudosos los incluimos
         if player[JSON_DATA][JSON_PLAYER_STATUS] == JSON_VALUE_OK || 
               ($options[:include_questionable_players] && player[JSON_DATA][JSON_PLAYER_STATUS] == JSON_VALUE_DOUBTFUL)
         
            f_data = Hash.new
            
            f_data[:player] = player
            f_data[:max_points] = player.get_max_points(1, @@current_week_number)
            f_data[:avg_points] = player.get_average_points(1, @@current_week_number)
            f_data[:last_points] = player.get_week_points(@@current_week_number)
            if !$options[:weeks].nil?
               f_data[:avg_week_points] = player.get_average_points($options[:weeks].first,$options[:weeks].last)
               f_data[:max_week_points] = player.get_max_points($options[:weeks].first,$options[:weeks].last)
               
               if f_data[:avg_week_points].nil?
                  f_data[:avg_week_points] = 0
               end
               
               if f_data[:max_week_points].nil?
                  f_data[:max_week_points] = 0
               end
            end
            
            if f_data[:max_points].nil?
               f_data[:max_points] = 0
            end
            
            if f_data[:avg_points].nil?
               f_data[:avg_points] = 0
            end
            
            if f_data[:last_points].nil?
               f_data[:last_points] = 0
            end
            
            case player[JSON_DATA][JSON_POSITION_ID]
               when 1
                  f_goal_keepers.push(f_data)
               when 2
                  f_defenders.push(f_data)
               when 3
                  f_midfielders.push(f_data)
               when 4
                  f_strikers.push(f_data)
            end
         end
      end
      
      return f_goal_keepers, f_defenders, f_midfielders, f_strikers
   end
   
   #
   # Crea la hoja que contiene las tablas dinámicas de la simulación por posiciones
   #
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_title [String] Nombre de la pestaña
   # @param p_goal_keepers [Array<Hash>] Array que contiene los porteros
   # @param p_defenders [Array<Hash>] Array que contiene los defensas
   # @param p_midfielders [Array<Hash>] Array que contiene los mediocampistas
   # @param p_strikers [Array<Hash>] Array que contiene los delanteros 
   #
   # @author Gerard Carrasquer
   #
   
   def write_simulation_sheet(p_workbook, p_title, p_goal_keepers, p_defenders, p_midfielders, p_strikers)
   
      #Text formats
      f_text_color = '#2e4053'
      f_text_font = 'Calibri'
      f_text_bg_color = '#fef9e7'
      
      # Formatos
      f_name_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
      
      f_worksheet = p_workbook.add_worksheet(p_title)
      
      # Establecemos el ancho de las columnas
   
      f_worksheet.set_column('B:B', 20) 
      f_worksheet.set_column('C:C', 18)
      f_worksheet.set_column('D:D', 18)
      f_worksheet.set_column('E:E', 18)
      if !$options[:weeks].nil?
         f_worksheet.set_column('F:F', 18)
         f_worksheet.set_column('G:G', 18)
      end
   
      # Establecemos el alto de las filas
   
      f_worksheet.set_row(1, 22)
      
      # Escritura de la hoja dinámica para la simulación de los porteros
      f_current_line, f_totals_line = write_simulation_table(p_workbook, f_worksheet, 4, p_goal_keepers, "Portero")
      
      f_max_points_cells = "C#{f_totals_line}" # Celdas para la suma de la puntuación máxima
      f_avg_points_cells = "D#{f_totals_line}" # Celdas para la suma de la puntuación media
      f_last_points_cells = "E#{f_totals_line}" # Celdas para la suma de la puntuación de la última jornada
      f_max_week_points_cells = nil # Celdas para la suma de la puntuación máxima de un rango de jornadas
      f_avg_week_points_cells = nil # Celdas para la suma de la puntuación media de un rango de jornadas
      
      #Se ha especificado un rango de jornadas. Damos valor a las variables de sus celdas
      if !$options[:weeks].nil?
         f_avg_week_points_cells = "F#{f_totals_line}"
         f_max_week_points_cells = "G#{f_totals_line}"
      end
      
      # Escritura de la hoja dinámica para la simulación de los defensas
      f_current_line, f_totals_line = write_simulation_table(p_workbook, f_worksheet, f_current_line, p_defenders, "Defensa")
      
      f_max_points_cells = "#{f_max_points_cells},C#{f_totals_line}"
      f_avg_points_cells = "#{f_avg_points_cells},D#{f_totals_line}"
      f_last_points_cells = "#{f_last_points_cells},E#{f_totals_line}"
      
      #Se ha especificado un rango de jornadas. Damos valor a las variables de sus celdas
      if !$options[:weeks].nil?
         f_avg_week_points_cells = "#{f_avg_week_points_cells},F#{f_totals_line}"
         f_max_week_points_cells = "#{f_max_week_points_cells},G#{f_totals_line}"
         
      end
      
      # Escritura de la hoja dinámica para la simulación de los mediocampistas
      f_current_line, f_totals_line = write_simulation_table(p_workbook, f_worksheet, f_current_line, p_midfielders, "Mediocampo")
      
      f_max_points_cells = "#{f_max_points_cells},C#{f_totals_line}"
      f_avg_points_cells = "#{f_avg_points_cells},D#{f_totals_line}"
      f_last_points_cells = "#{f_last_points_cells},E#{f_totals_line}"
      
      #Se ha especificado un rango de jornadas. Damos valor a las variables de sus celdas
      if !$options[:weeks].nil?
         f_avg_week_points_cells = "#{f_avg_week_points_cells},F#{f_totals_line}"
         f_max_week_points_cells = "#{f_max_week_points_cells},G#{f_totals_line}"
      end
      
      # Escritura de la hoja dinámica para la simulación de los delanteros
      f_current_line, f_totals_line = write_simulation_table(p_workbook, f_worksheet, f_current_line, p_strikers, "Delantera")
      
      f_max_points_cells = "#{f_max_points_cells},C#{f_totals_line}"
      f_avg_points_cells = "#{f_avg_points_cells},D#{f_totals_line}"
      f_last_points_cells = "#{f_last_points_cells},E#{f_totals_line}"
      
      #Se ha especificado un rango de jornadas. Damos valor a las variables de sus celdas
      if !$options[:weeks].nil?
         f_avg_week_points_cells = "#{f_avg_week_points_cells},F#{f_totals_line}"
         f_max_week_points_cells = "#{f_max_week_points_cells},G#{f_totals_line}"
      end
      
      # Escritura de las fórmulas de los totales de puntuación de cada criterio de puntuación 
      f_worksheet.write("B2", 'Puntos totales', f_name_format)
      f_worksheet.write("C2", "=SUM(#{f_max_points_cells})", f_name_format)
      f_worksheet.write("D2", "=SUM(#{f_avg_points_cells})", f_name_format)
      f_worksheet.write("E2", "=SUM(#{f_last_points_cells})", f_name_format)
      
      #Se ha especificado un rango de jornadas. Escribimos sus fórmulas 
      if !$options[:weeks].nil?
         f_worksheet.write("F2", "=SUM(#{f_avg_week_points_cells})", f_name_format)
         f_worksheet.write("G2", "=SUM(#{f_max_week_points_cells})", f_name_format)
      end
   end
   
   #
   # Crea la hoja que contiene el cálculo de la alineación óptima
   #
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_title [String] Nombre de la pestaña
   # @param p_goal_keepers [Array<Hash>] Array que contiene los porteros
   # @param p_defenders [Array<Hash>] Array que contiene los defensas
   # @param p_midfielders [Array<Hash>] Array que contiene los mediocampistas
   # @param p_strikers [Array<Hash>] Array que contiene los delanteros 
   #
   # @author Gerard Carrasquer
   #
   def write_lineup_sheet(p_workbook, p_title, p_goal_keepers, p_defenders, p_midfielders, p_strikers)
      
      f_worksheet = p_workbook.add_worksheet(p_title)
      
      f_worksheet.set_column('B:B', 47)
      f_worksheet.set_column('E:E', 45) 
      
      # Escritura de la alineación óptima según la puntuación máxima de cada jugador
      f_current_line = write_lineup(p_workbook, f_worksheet, 2, 'B', :max_points, 'Puntuación máxima', 
                 p_goal_keepers, p_defenders, p_midfielders, p_strikers)
      
      # Escritura de la alineación óptima según la puntuación media de cada jugador           
      f_current_line = write_lineup(p_workbook, f_worksheet, 2, 'E', :avg_points, 'Puntuación media', 
                 p_goal_keepers, p_defenders, p_midfielders, p_strikers)
                 
      # Escritura de la alineación óptima según la puntuación de la última jornada de cada jugador
      write_lineup(p_workbook, f_worksheet, f_current_line, 'B', :last_points, 'Puntuación última jornada', 
                 p_goal_keepers, p_defenders, p_midfielders, p_strikers)
      
      #Se ha especificado un rango de jornadas.
      if !$options[:weeks].nil?
         # Escritura de la alineación óptima según la puntuación media en el rango de jornadas de cada jugador
         f_current_line = write_lineup(p_workbook, f_worksheet, f_current_line, 'E', :avg_week_points, 
                  "Puntuación media de las jornadas #{$options[:weeks].first} a #{$options[:weeks].last}", 
                 p_goal_keepers, p_defenders, p_midfielders, p_strikers)
         # Escritura de la alineación óptima según la puntuación máxima en el rango de jornadas de cada jugador
         f_current_line = write_lineup(p_workbook, f_worksheet, f_current_line, 'B', :max_week_points, 
                  "Puntuación máxima de las jornadas #{$options[:weeks].first} a #{$options[:weeks].last}", 
                 p_goal_keepers, p_defenders, p_midfielders, p_strikers)
      end
   end
   
   #
   # Creación de la tabla dinámica de una posición para la simulación de alineaciones.
   #
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   # @param p_line [Integer] Línea de la tabla
   # @param p_pos_players [Array<Hash>] Jugadores de la posición
   # @param p_title [String] Nombre de la posición 
   #
   # @return [Integer] Línea final de las informaciones escritas.
   # @return [String] Celda que contiene la fórmula del total de puntos de la posición.
   #
   # @author Gerard Carrasquer
   #
   def write_simulation_table(p_workbook, p_worksheet, p_line, p_pos_players, p_title)
   
      f_current_line = p_line
      
      #Text formats
      f_text_color = '#2e4053'
      f_text_bg_color = '#fef9e7'
      f_totals_bg_color = '#f5b7b1'

      f_info_color = '#34495e'
      f_doubtful_bg_color = '#f9e79f'
      f_header_color = '#34495e'
      f_header_bg_color = '#eafaf1'
      f_info_bg_color = '#ebf5fb'
      f_text_font = 'Calibri'
      
      # Formatos
      f_name_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
      f_info_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_info_bg_color, :border => 1)
      f_doubtful_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_doubtful_bg_color, :border => 1)
      f_totals_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_totals_bg_color, :border => 1, :bold => 1)
      f_header_format = p_workbook.add_format(:color => f_header_color, :size => 10, :font => f_text_font, 
                        :bg_color => f_header_bg_color, :border => 1, :bold => 1, :align => 'justify')
      f_currency_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font, 
                        :num_format => '#,##0', :bg_color => f_info_bg_color, :border => 1)
      
      p_worksheet.set_row(f_current_line - 1, 22)
      
      p_worksheet.merge_range("B#{f_current_line}:F#{f_current_line}", p_title, f_name_format)
      
      f_current_line += 2
      
      p_worksheet.set_row(f_current_line - 1, 27)
      
      f_table_last_column ='E'
      f_columns = Array.new
      
      #Cabeceras de las tablas dinámicas con su autofilter
      f_columns.push({ :header => 'Nombre', :header_format => f_header_format })
      f_columns.push({ :header => 'Puntuación máxima', :header_format => f_header_format})
      f_columns.push({ :header => 'Puntuación media', :header_format => f_header_format})
      f_columns.push({ :header => 'Puntuación última jornada', :header_format => f_header_format})
      
      #Se ha especificado un rango de jornadas. Añadimos las cabeceras de las columnas correspondientes
      if !$options[:weeks].nil?
         f_columns.push({ :header => "Puntuación media jorn. #{$options[:weeks].first} a #{$options[:weeks].last}", 
               :header_format => f_header_format})
         f_columns.push({ :header => "Puntuación máxima jorn. #{$options[:weeks].first} a #{$options[:weeks].last}", 
               :header_format => f_header_format})
         f_table_last_column ='G'
      end
      
      p_worksheet.add_table("B#{f_current_line}:#{f_table_last_column}#{f_current_line + p_pos_players.length}", :columns => f_columns,
                    :style     => 'Table Style Light 11',
                    :name => p_title)
            
      f_current_line += 1
      
      # Fórmulas sobre los elementos seleccionados en las tablas dinámicas.
      # 109 => Suma de las celdas seleccionadas dentro del rango
      # 103 => Cuenta el número de celdas seleccionadas dentro del rango
      p_worksheet.write("C#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(109,C#{f_current_line}:C#{f_current_line + p_pos_players.length - 1})", f_totals_format)
      p_worksheet.write("D#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(109,D#{f_current_line}:D#{f_current_line + p_pos_players.length - 1})", f_totals_format)
      p_worksheet.write("E#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(109,E#{f_current_line}:E#{f_current_line + p_pos_players.length - 1})", f_totals_format)
      p_worksheet.write("B#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(103,B#{f_current_line}:B#{f_current_line + p_pos_players.length - 1})", f_totals_format)
      if !$options[:weeks].nil?
         p_worksheet.write("F#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(109,F#{f_current_line}:F#{f_current_line + p_pos_players.length - 1})", f_totals_format)
         p_worksheet.write("G#{f_current_line + p_pos_players.length + 1}", "=SUBTOTAL(109,G#{f_current_line}:G#{f_current_line + p_pos_players.length - 1})", f_totals_format)
      end
      
      f_totals_line = f_current_line + p_pos_players.length + 1
      
      p_pos_players.each do |p|
         
         f_format = f_info_format
         if p[:player][JSON_DATA][JSON_PLAYER_STATUS] == JSON_VALUE_DOUBTFUL
            # Los jugadores dudosos tienen un color diferente
            f_format = f_doubtful_format
         end
         
         p_worksheet.write("B#{f_current_line}", p[:player][JSON_DATA][JSON_NICKNAME], f_format)
         p_worksheet.write_number("C#{f_current_line}", p[:max_points], f_format)
         p_worksheet.write_number("D#{f_current_line}", p[:avg_points], f_format)
         p_worksheet.write_number("E#{f_current_line}", p[:last_points], f_format)
         
         if !$options[:weeks].nil?
            p_worksheet.write_number("F#{f_current_line}", p[:avg_week_points], f_format)
            p_worksheet.write_number("G#{f_current_line}", p[:max_week_points], f_format)
         end
         f_current_line += 1
      end
            
      f_current_line += 3
      
      return f_current_line, f_totals_line
   end
   
   #
   # Escribe la alineación óptima en la hoja
   #
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   # @param p_line [Integer] Línea en la que se empezará a escribir
   # @param p_column [String] Columna en la que se empezará a escribir
   # @param p_calc_criteria [Symbol] Criterio de puntuación. Corresponde a la clave de los Hash dentro de los 
   #        arrays de jugadores que almacena los puntos 
   # @param p_title [String] Nombre de la alineación
   # @param p_goalkeepers [Array<Hash>] Array que contiene los porteros
   # @param p_defenders [Array<Hash>] Array que contiene los defensas
   # @param p_midfielders [Array<Hash>] Array que contiene los mediocampistas
   # @param p_strikers [Array<Hash>] Array que contiene los delanteros 
   #
   # @author Gerard Carrasquer
   #
   def write_lineup(p_workbook, p_worksheet, p_line, p_column, p_calc_criteria, p_title, 
                    p_goalkeepers, p_defenders, p_midfielders, p_strikers)
                    
      f_current_line = p_line
      
      #Text formats
      f_text_color = '#2e4053'
      f_text_bg_color = '#fef9e7'
      f_totals_bg_color = '#f5b7b1'

      f_info_color = '#34495e'
      f_header_color = '#34495e'
      f_header_bg_color = '#eafaf1'
      f_info_bg_color = '#ebf5fb'
      f_text_font = 'Calibri'
      
      # Formatos
      f_name_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
      f_subtitle_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 12, :font => f_text_font,
                        :bg_color => f_text_bg_color, :border => 1)
      f_info_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_info_bg_color, :border => 1)
      f_totals_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font,
                        :bg_color => f_totals_bg_color, :border => 1, :bold => 1)
      f_header_format = p_workbook.add_format(:color => f_header_color, :size => 10, :font => f_text_font, 
                        :bg_color => f_header_bg_color, :border => 1, :bold => 1, :align => 'justify')
      f_currency_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font, 
                        :num_format => '#,##0', :bg_color => f_info_bg_color, :border => 1)
      
      p_worksheet.set_row(f_current_line - 1, 22)
      
      # Posibles alineaciones
      f_lineups = [[5,4,1], [5,3,2], [4,5,1], [4,4,2], [4,3,3], [3,5,2], [3,4,3]]
         
      # Ordenamos los arrays de jugadores con el criterio seleccionado de forma descendente. El criterio seleccionado
      # es la clave del hash de cada jugador dentro de cada array
      p_goalkeepers.sort_by!{|player| -1 * player[p_calc_criteria]}
      p_defenders.sort_by!{|player| -1 * player[p_calc_criteria]}
      p_midfielders.sort_by!{|player| -1 * player[p_calc_criteria]}
      p_strikers.sort_by!{|player| -1 * player[p_calc_criteria]}
         
      p_worksheet.merge_range("#{p_column}#{f_current_line}:#{p_column.succ}#{f_current_line}", p_title, f_name_format)
         
      f_points = 0
      f_position = nil
      
       
      # Recorremos las alineaciones para encontrar la que maximiza los puntos según el criterio seleccionado 
      f_lineups.each do |line|
         
         if p_goalkeepers.empty?
            puts "Porteros insuficientes para la alineación #{line.join('-')}"
         elsif p_defenders.length < line[0]
            puts "Defensas insuficientes para la alineación #{line.join('-')}"
         elsif p_midfielders.length < line[1]
            puts "Mediocampistas insuficientes para la alineación #{line.join('-')}"
         elsif p_strikers.length < line[2]
            puts "Delanteros insuficientes para la alineación #{line.join('-')}"
         else
            # Cálculo de las alineación óptima
            # Los arrays de jugadodres están ordenados de forma descendente. Cogemos los n primeros hasta
            # completar todos los jugadores disponibles posiciones  
            f_current_points = p_goalkeepers[0][p_calc_criteria]
            
            # Defensa
            0.upto(line[0] - 1) do |i|
               f_current_points += p_defenders[i][p_calc_criteria]
            end
            
            # Mediocampo
            0.upto(line[1] - 1) do |i|
               f_current_points += p_midfielders[i][p_calc_criteria]
            end
            
            # Delantera
            0.upto(line[2] - 1) do |i|
               f_current_points += p_strikers[i][p_calc_criteria]
            end
            
            # Si la puntuación total es mayor que la almacenada previamente lo actualizamos
            if f_current_points > f_points
               f_points = f_current_points
               f_position = line
            end
         end
      end
      f_current_line += 1 
      
      # Escritura en el excel
      p_worksheet.merge_range("#{p_column}#{f_current_line}:#{p_column.succ}#{f_current_line}", f_position.join('-'), f_subtitle_format)
      
      f_current_line += 1
      f_sum_cells = "#{p_column.succ}#{f_current_line}"
      
      p_worksheet.write("#{p_column}#{f_current_line}", p_goalkeepers[0][:player][JSON_DATA][JSON_NICKNAME], f_info_format)
      p_worksheet.write_number("#{p_column.succ}#{f_current_line}", p_goalkeepers[0][p_calc_criteria], f_info_format)
      
      f_current_line += 2
      
      # Defensa
      0.upto(f_position[0] - 1) do |i|
         p_worksheet.write("#{p_column}#{f_current_line}", p_defenders[i][:player][JSON_DATA][JSON_NICKNAME], f_info_format)
         p_worksheet.write_number("#{p_column.succ}#{f_current_line}", p_defenders[i][p_calc_criteria], f_info_format)
         
         f_sum_cells = "#{f_sum_cells}+#{p_column.succ}#{f_current_line}"
         f_current_line += 1
      end
      
      f_current_line += 1
      
      # Mediocampo
      0.upto(f_position[1] - 1) do |i|
         p_worksheet.write("#{p_column}#{f_current_line}", p_midfielders[i][:player][JSON_DATA][JSON_NICKNAME], f_info_format)
         p_worksheet.write_number("#{p_column.succ}#{f_current_line}", p_midfielders[i][p_calc_criteria], f_info_format)
         
         f_sum_cells = "#{f_sum_cells}+#{p_column.succ}#{f_current_line}"
         f_current_line += 1
      end
      
      f_current_line += 1
      
      # Delantera
      0.upto(f_position[2] - 1) do |i|
         p_worksheet.write("#{p_column}#{f_current_line}", p_strikers[i][:player][JSON_DATA][JSON_NICKNAME], f_info_format)
         p_worksheet.write_number("#{p_column.succ}#{f_current_line}", p_strikers[i][p_calc_criteria], f_info_format)
         
         f_sum_cells = "#{f_sum_cells}+#{p_column.succ}#{f_current_line}"
         f_current_line += 1
      end
      
      p_worksheet.write_formula("#{p_column.succ}#{f_current_line}", "=#{f_sum_cells}", f_totals_format)
      
      f_current_line += 2
      
      return f_current_line
      
   end
end
