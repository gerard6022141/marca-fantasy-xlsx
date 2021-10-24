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
         if (!$options[:players_file].nil? && $options[:compare_players]) || !$options[:players_folder].nil?
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
                        JSON_PENALTY_FAILED => "Penaltis falados", 
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
                        JSON_MARCA_POINTS => "Puntos marca"}
         
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
                                 :init_line => f_current_row, :end_line => f_current_row + f_weeks_count})
                     end
                     
                     # Obtenemos las estadísticas del jugador para una jornada determinada
                     f_player_stats = p.get_week_stats(week)
                     
                     
                     if f_player_stats.nil?
                        # La jornada especificada no está en el JSON que hemos descargado, lo cual significa
                        # que el jugador no ha jugado esa jornada. Rellenamos la celda correspondiente a esa jornada
                        # con '-''
                        f_worksheet.write("#{f_current_column}#{f_current_row}", '-',
                                          (f_current_row.odd?) ? f_odd_format : f_even_format)
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
                  'column', "C#{f_chart_row + (i * 22)}", v, 'Jornada', v, *f_series)
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
      f_filter_ok = $options[:players_id].nil? && $options[:players_names].nil?
       
      if !$options[:players_id].nil?
         f_filter_ok = $options[:players_id].include?(p_player_id)
      end
      
      if !$options[:players_names].nil?
         f_filter_ok |= !$options[:players_names].index{|e| 
                  Regexp.new(".*#{I18n.transliterate(e)}.*", true).match?(I18n.transliterate(p_player_name))}.nil?
      end
      
      
      
      return f_filter_ok
   end
end
