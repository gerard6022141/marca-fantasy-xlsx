require './fantasy_constants.rb'
require './fantasy.rb'
require 'write_xlsx'
require 'down'

#
# Clase que contiene la información estadística de un jugador
#

class Player < Fantasy

   #
   # Constructor. Descarga el JSON que contiene la información estadística de un jugador.
   # @param p_player_id [Integer] Identificador del jugador 
   #
   # @raise [Timeout::Error, Errno::EINVAL, Errno::ECONNRESET, EOFError,
   #          Net::HTTPBadResponse, Net::HTTPHeaderSyntaxError, Net::ProtocolError] Error en la descarga del fichero JSON
   # @author Gerard Carrasquer
   #
   
   def initialize(p_player_id)
      
      @player_data = Hash.new    # JSON que contiene la inforamción estadística del jugador
      @stats_columns = Hash.new  # Columnas de la estadística
      
      super()
      
      @player_data = get_data("#{FANTASY_API_SERVER}/#{FANTASY_PLAYER_URL}#{p_player_id}")
            
   end
   
   #
   # Devuelve el dato correpondiente a un tag en el JSON que contiene la información estadística de un jugador.
   # @param p_json_tag [String] Nombre del TAG 
   #
   # @return [String, Integer, Float, Bool, nil]
   # @author Gerard Carrasquer
   #
   
   def [](p_json_tag)
      return @player_data[p_json_tag]
   end
   
   #
   # Crea un fichero excel a partir del JSON que contiene la información estadística de un jugador.
   # @param p_file_name [String] Nombre del fichero 
   #
   # @raise [Exception] Error en la escritura del fichero excel
   #
   # @author Gerard Carrasquer
   #
   
   def to_xlsx(p_file_name)
      f_ok = true
      if File.exist?(p_file_name)
         begin
            File.delete(p_file_name)
         rescue => e
            f_ok = false
            puts e.to_s
            if $options[:verbose]
               puts "#{e.to_s}\n\n#{e.backtrace}"
            end
         end
      end
      
      if f_ok
         print "Generando fichero #{p_file_name}... "
            
         f_workbook = WriteXLSX.new(p_file_name)

         f_worksheet = f_workbook.add_worksheet('Estadísticas')
         
         player_head_to_xlsx(f_workbook, f_worksheet)
         
         player_stats_to_xlsx(f_workbook, f_worksheet)
         
         if $options[:chart]
            # Generamos la pestaña de gráficos si se ha especificado en línea de comandos
            f_points_chart_worksheet = f_workbook.add_worksheet('Gráficas')
         
            f_weeks_count = @@current_week_number
            f_first_week = 1
            if !$options[:weeks].nil? 
               f_weeks_count = $options[:weeks].length
               f_first_week = $options[:weeks][0]
            end
            
            @stats_columns.each_with_index do |(k, v), i|
            
               # Generamos un gráfico por cada dato estadístico del jugador
               if i != 0 && i!= 21
                  # Todos los datos estadísticos tiene asociados los puntos otorgados excepto los puntos
                  # totales del jugador en la jornada y los puntos marca (índices 0 y 21 del array)
                  # Hacemos la llamada a la función con dos series: una para el dato estadístico y otra
                  # para los puntos otorgados
                  draw_column_chart(f_workbook, f_points_chart_worksheet, f_worksheet.name, 
                     'column', "B#{(30 * i) + 2}", v, 'Jornada', v, 
                     {:name => v, :categories_column => 'C', :values_column => k, :init_line => 18, :end_line => f_weeks_count + 17},
                     {:name => 'Puntos', :categories_column => 'D', :values_column => k, :init_line => 18 + f_weeks_count + 7, :end_line => (2 * f_weeks_count) + 18 + 6})
               else
                  #Gráfico para una sola serie
                  draw_column_chart(f_workbook, f_points_chart_worksheet, f_worksheet.name, 
                     'column', "B#{(30 * i) + 2}", v, 'Jornada', v, 
                     {:name => v, :categories_column => 'C', :values_column => k, :init_line => 18, :end_line => f_weeks_count + 17})
               end
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
   end
   
   #
   # Escribe la información de la cabecera del jugador en el fichero excel.
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   #
   # @author Gerard Carrasquer
   #
   
   def player_head_to_xlsx(p_workbook, p_worksheet)
   
      #Text formats
      f_text_color = '#d0d3d4'
      f_info_color = '#839192'
      f_text_font = 'Calibri'
      f_name_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font)
      f_info_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font)
      f_currency_format = p_workbook.add_format(:color => f_info_color, :size => 10, :font => f_text_font, :num_format => '#,##0')
      

      draw_spreadsheet_border(p_workbook, p_worksheet, '#3498db', 'B2', 'K12')
      
      # Descarga de la foto del jugador y el escudo del equipo
      f_tempfile = Down.download(@player_data[JSON_DATA][JSON_IMAGES][JSON_TRANSPARENT][JSON_128X128])
      p_worksheet.insert_image(2, 1, f_tempfile)
      
      f_tempfile = Down.download(@player_data[JSON_DATA][JSON_TEAM][JSON_BADGE_COLOR])
      p_worksheet.insert_image(2, 9, f_tempfile)
      
      
      p_worksheet.set_row(3, 20)
      
      #Nombre
      p_worksheet.merge_range('D4:I4', @player_data[JSON_DATA][JSON_NAME], f_name_format)
      
      #Posición
      p_worksheet.merge_range('D5:F5', 'Posición', f_info_format)
      p_worksheet.write('G5', @player_data[JSON_DATA][JSON_POSITION], f_info_format)
      
      #Estado
      p_worksheet.merge_range('D6:F6', 'Estado', f_info_format)
      p_worksheet.write('G6', @player_data[JSON_DATA][JSON_PLAYER_STATUS], f_info_format)
      
      #Puntos en la última temporada
      p_worksheet.merge_range('D7:F7', 'Puntos en la última temporada', f_info_format)
      if @player_data[JSON_DATA][JSON_LAST_SEASON_POINTS].nil?
         p_worksheet.write('G7', '.', f_info_format)
      else
         p_worksheet.write_number('G7', @player_data[JSON_DATA][JSON_LAST_SEASON_POINTS], f_info_format)
      end
      
      #Puntos en la temporada actual
      p_worksheet.merge_range('D8:F8', 'Puntos en la temporada actual', f_info_format)
      p_worksheet.write_number('G8', @player_data[JSON_DATA][JSON_POINTS], f_info_format)
      
      
      #Media de puntos
      p_worksheet.merge_range('D9:F9', 'Media de puntos', f_info_format)
      p_worksheet.write_number('G9', @player_data[JSON_DATA][JSON_AVERAGE_POINTS].round(2), f_info_format)
      
      #Valor de mercado
      p_worksheet.merge_range('D10:F10', 'Valor de mercado', f_info_format)
      p_worksheet.write_number('G10', @player_data[JSON_DATA][JSON_MARKET_VALUE], f_currency_format)
      
      
      
   end
   
   #
   # Escribe la información estadística del jugador en el fichero excel.
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   #
   # @author Gerard Carrasquer
   #
   
   def player_stats_to_xlsx(p_workbook, p_worksheet)
   
      #Text formats
      f_text_color = '#34495e'
      f_odd_color = '#fcf3cf'
      f_even_color = '#d5f5e3'
      f_text_font = 'Calibri'
      f_odd_format = p_workbook.add_format(:bold => 0, :color => f_text_color, :size => 10,
                  :font => f_text_font, :bg_color => f_odd_color, :border => 1)
      f_even_format = p_workbook.add_format(:bold => 0, :color => f_text_color, :size => 10,
                  :font => f_text_font, :bg_color => f_even_color, :border => 1)
      f_tag_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 10, :font => f_text_font,
                  :border => 1)
      f_title_format = p_workbook.add_format(:bold => 1, :color => f_text_color, :size => 16, :font => f_text_font)
      
      f_weeks_count = @@current_week_number
      f_first_week = 1
      if !$options[:weeks].nil? 
         f_weeks_count = $options[:weeks].length
         f_first_week = $options[:weeks][0]
      end
      
      if not @@current_week_number.nil?
      
         # La información de la jorada en curso se ha descargado correctamente
         
         f_points_line = 16 + f_weeks_count + 6
         #Title
         p_worksheet.set_row(14, 20)
         p_worksheet.set_row(16, 35)
         p_worksheet.set_row(f_points_line - 1, 20)
         p_worksheet.set_row(f_points_line + 1, 35)
         p_worksheet.set_column('D:Y', 10)
         p_worksheet.set_column('C:C', 15)
         
         p_worksheet.merge_range('B15:F15', 'Estadísticas', f_title_format)
         
         draw_spreadsheet_border(p_workbook, p_worksheet, '#3498db', 'B16', "Z#{16 + f_weeks_count + 3}")
         
         p_worksheet.merge_range("B#{f_points_line}:F#{f_points_line}", 
               'Puntos', f_title_format)
         
         draw_spreadsheet_border(p_workbook, p_worksheet, '#3498db', "B#{f_points_line + 1}", "Y#{f_points_line + f_weeks_count + 3}")
         
         # Ordenamos el array que contiene las estadísticas de la jornada por el número de jornada
         # porque a veces viene desordenado
         @player_data[JSON_DATA][JSON_PLAYER_STATS].sort_by!{|week| week[JSON_WEEK_NUMBER]}
         
         
         f_column_names_array = ["Puntos", 
               "Minutos\njugados", 
               "Goles", 
               "Asistencias\nde gol",
               "Asistencias\nsin gol", 
               "Balones\nal área", 
               "Penaltis\nprovocados", 
               "Penaltis\nparados",
               "Paradas", 
               "Despejes", 
               "Penaltis\nfallados", 
               "Goles en\npropia\npuerta", 
               "Goles\nen contra", 
               "Tarjeta\namarilla", 
               "Segunda\ntarjeta\namarilla", 
               "Tarjeta\nroja", 
               "Tiros\na puerta",
               "Regates", 
               "Balones\nrecuperados", 
               "Posesiones\nperdidas", 
               "Penaltis\ncometidos", 
               "Puntos\nmarca"]
               
         f_stat_column = 'D'
         
         # Añadimos los nombres de las columnas al array de nombres eliminando los saltos de línea
         f_column_names_array.each do |c|
            @stats_columns[f_stat_column] = c.gsub(/\n/, ' ')
            f_stat_column.succ!
         end
         
         f_week = 1 # Número de jornada tratana en cada iteración del each de @player_data
         f_current_row = 16
         p_worksheet.write_row(f_current_row, 3, f_column_names_array, f_tag_format)
         p_worksheet.write_row(f_points_line + 1, 3, f_column_names_array[0 .. 20], f_tag_format)
         
         f_current_row += 1
         
         @player_data[JSON_DATA][JSON_PLAYER_STATS].each do |v|
            
            # Recorremos el array de estadísticas. Contiene un JSON por cada jornada jugada
            
            
            
            if f_week != v[JSON_WEEK_NUMBER]
               # Las jornadas en las que jugador no ha jugado no están contenidas en el array de estadísticas del
               # jugador. Si la jornada que estamos tratando es menor de la jornada contenida en el fichero escribimos
               # una línea en blanco para cada jornada que falte hasta llegar a la que contiene el array  
               f_week.upto(v[JSON_WEEK_NUMBER] - 1) do |week|
               
                  if $options[:weeks].nil? || $options[:weeks].include?(week)
                     # Si se ha especificado la opción --weeks en línea de comandos $options[:compare_weeks]
                     # contiene las jornadas a comparar especificadas por el usuario.
                     # En caso contrario $options[:compare_weeks] es nil y debemos recorrer desde la jornada 1
                     # hasta la jornada actual
                     #Líneas de estadísticas
                     p_worksheet.write("C#{f_current_row + 1}", "Jornada #{week}", f_tag_format)
                     p_worksheet.write_row(f_current_row, 3, Array.new(22, '-'),
                                    (week.odd?) ? f_odd_format : f_even_format)
                     #Líneas de puntos
                     p_worksheet.write("C#{f_current_row + f_weeks_count + 8}", "Jornada #{week}", f_tag_format)
                     p_worksheet.write_row(f_current_row + f_weeks_count + 7, 3, Array.new(21, '-'),
                                    (week.odd?) ? f_odd_format : f_even_format)
                     
                     f_current_row += 1
                  end
               end
            end
            
            if $options[:weeks].nil? || $options[:weeks].include?(v[JSON_WEEK_NUMBER])
               # Si se ha especificado la opción --weeks en línea de comandos $options[:compare_weeks]
               # contiene las jornadas a comparar especificadas por el usuario.
               # En caso contrario $options[:compare_weeks] es nil y debemos recorrer desde la jornada 1
               # hasta la jornada actual
               
               #línea de estadísticas
               p_worksheet.write("C#{f_current_row + 1}", "Jornada #{v[JSON_WEEK_NUMBER]}", f_tag_format)
               
               # Cada elemento del array de estadísticas contiene un array de dos elementos: el cero es el dato
               # estadístico y el uno los puntos conseguidos con ese dato estadístico. Escribimos uno en cada tabla
               # Los puntos totales están en la raiz del JSON
               # Los puntos marca están en el elemento uno del array
               p_worksheet.write_row(f_current_row, 3, [v[JSON_TOTAL_POINTS], 
                           v[JSON_STATS][JSON_MINS_PLAYED][0],
                           v[JSON_STATS][JSON_GOALS][0], 
                           v[JSON_STATS][JSON_GOAL_ASSIST][0], 
                           v[JSON_STATS][JSON_GOAL_OFFTARGET_ATT_ASSIST][0], 
                           v[JSON_STATS][JSON_PEN_AREA_ENTRIES][0], 
                           v[JSON_STATS][JSON_PENALTY_WON][0], 
                           v[JSON_STATS][JSON_PENALTY_SAVE][0],
                           v[JSON_STATS][JSON_SAVES][0], 
                           v[JSON_STATS][JSON_EFECTIVE_CLEARANCE][0], 
                           v[JSON_STATS][JSON_PENALTY_FAILED][0], 
                           v[JSON_STATS][JSON_OWN_GOALS][0], 
                           v[JSON_STATS][JSON_GOALS_CONCEDED][0], 
                           v[JSON_STATS][JSON_YELLOW_CARD][0], 
                           v[JSON_STATS][JSON_SECOND_YELLOW_CARD][0], 
                           v[JSON_STATS][JSON_RED_CARD][0], 
                           v[JSON_STATS][JSON_TOTAL_SCORING_ATT][0], 
                           v[JSON_STATS][JSON_WON_CONTEST][0], 
                           v[JSON_STATS][JSON_BALL_RECOVERY][0], 
                           v[JSON_STATS][JSON_POSS_LOST_ALL][0], 
                           v[JSON_STATS][JSON_PENALTY_CONCEDED][0], 
                           v[JSON_STATS][JSON_MARCA_POINTS][1]], 
                           (v[JSON_WEEK_NUMBER].odd?) ? f_odd_format : f_even_format)
               
               #Línea de puntos
               p_worksheet.write("C#{f_current_row + f_weeks_count + 8}", "Jornada #{v[JSON_WEEK_NUMBER]}", f_tag_format)
               p_worksheet.write_row(f_current_row + f_weeks_count + 7, 3, ['-', 
                           v[JSON_STATS][JSON_MINS_PLAYED][1],
                           v[JSON_STATS][JSON_GOALS][1], 
                           v[JSON_STATS][JSON_GOAL_ASSIST][1], 
                           v[JSON_STATS][JSON_GOAL_OFFTARGET_ATT_ASSIST][1], 
                           v[JSON_STATS][JSON_PEN_AREA_ENTRIES][1], 
                           v[JSON_STATS][JSON_PENALTY_WON][1], 
                           v[JSON_STATS][JSON_PENALTY_SAVE][1],
                           v[JSON_STATS][JSON_SAVES][1], 
                           v[JSON_STATS][JSON_EFECTIVE_CLEARANCE][1], 
                           v[JSON_STATS][JSON_PENALTY_FAILED][1], 
                           v[JSON_STATS][JSON_OWN_GOALS][1], 
                           v[JSON_STATS][JSON_GOALS_CONCEDED][1], 
                           v[JSON_STATS][JSON_YELLOW_CARD][1], 
                           v[JSON_STATS][JSON_SECOND_YELLOW_CARD][1], 
                           v[JSON_STATS][JSON_RED_CARD][1], 
                           v[JSON_STATS][JSON_TOTAL_SCORING_ATT][1], 
                           v[JSON_STATS][JSON_WON_CONTEST][1], 
                           v[JSON_STATS][JSON_BALL_RECOVERY][1], 
                           v[JSON_STATS][JSON_POSS_LOST_ALL][1], 
                           v[JSON_STATS][JSON_PENALTY_CONCEDED][1]], 
                           (v[JSON_WEEK_NUMBER].odd?) ? f_odd_format : f_even_format)
                           
               f_current_row += 1
               
            end            
            f_week = v[JSON_WEEK_NUMBER] + 1
               
         end
         
         # Las jornadas en las que jugador no ha jugado no están contenidas en el array de estadísticas del
         # jugador. Si salimos del bucle en una jornada inferior a la jornada actual escribimos
         # una línea en blanco para cada jornada que falte hasta llegar a la jornada actual
         f_week.upto(@@current_week_number) do |week|
            if $options[:weeks].nil? || $options[:weeks].include?(week)
               # Si se ha especificado la opción --weeks en línea de comandos $options[:compare_weeks]
               # contiene las jornadas a comparar especificadas por el usuario.
               # En caso contrario $options[:compare_weeks] es nil y debemos recorrer desde la jornada 1
               # hasta la jornada actual
               
               f_current_row += 1
               
               #Línea de estadísticas
               p_worksheet.write("C#{f_current_row + 1}", "Jornada #{week}", f_tag_format)
               p_worksheet.write_row(f_current_row, 3, Array.new(22, '-'),
                              (week.odd?) ? f_odd_format : f_even_format)
                              
               #Línea de puntos
               p_worksheet.write("C#{f_current_row + @@current_week_number + 8}", "Jornada #{week}", f_tag_format)
               p_worksheet.write_row(f_current_row + @@current_week_number + 7, 3, Array.new(21, '-'),
                              (week.odd?) ? f_odd_format : f_even_format)
            end
         end
      end
      
   end
   
   #
   # Devuelve la información estadística de una jornada.
   # @param p_week_number [Integer] Número de jornada 
   #
   # @return [Hash, nil]
   # @author Gerard Carrasquer
   #
   
   def get_week_stats(p_week_number)
      
      @player_data[JSON_DATA][JSON_PLAYER_STATS].each do |v|
         if v[JSON_WEEK_NUMBER] == p_week_number
            return v
         end
      end 
      
      return nil
   end
   

end
