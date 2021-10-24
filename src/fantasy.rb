require 'uri'
require 'net/http'
require 'json'
require './fantasy_constants.rb'

#
# Clase base para los jugadores y la lista de jugadores
#

class Fantasy

   # Jornada actual
   @@current_week_number = nil

   # Constructor
   
   def initialize
      if @@current_week_number.nil?
         # Descarga de la jornada actual. Se descarga solo una vez porque siempre va a ser la misma
         @@current_week_number = get_current_week
      end
   end

   #
   # Descarga de un archivo JSON desde una URI.
   #
   # @param p_uri [String] URI desde la que se descargará el fichero
   # @return [Hash] Una variable de tipo Hash que contiene el JSON descargado.
   #                Los datos descargados están en el tag "data" dentro del JSON
   #                El resultado de la descarga se encuentra en el tag "response" ("ok" o "error")
   #                En caso de error el tag "message" contiene una descripción
   # @raise [Timeout::Error, Errno::EINVAL, Errno::ECONNRESET, EOFError,
   #          Net::HTTPBadResponse, Net::HTTPHeaderSyntaxError, Net::ProtocolError] Error en la descarga del fichero JSON
   # @author Gerard Carrasquer
   #
   def get_data(p_uri)
      f_data = Hash.new
      f_data[JSON_DATA] = Hash.new
      begin
         f_uri = URI(p_uri)
         f_res = Net::HTTP.get_response(f_uri)
         
         if $options[:verbose]
            puts "Conectando a #{f_uri.inspect} Respuesta #{f_res.code}"
         end
      
         if f_res.code.to_i != 200
            
            f_data[JSON_RESPONSE] = JSON_ERROR
            f_data[JSON_DATA][JSON_MESSAGE] = "Error #{f_res.code} - #{f_res.message}"
         else
            f_data[JSON_RESPONSE] = JSON_OK
            f_data[JSON_DATA] = JSON.parse(f_res.body)
         end
      rescue Timeout::Error, Errno::EINVAL, Errno::ECONNRESET, EOFError,
             Net::HTTPBadResponse, Net::HTTPHeaderSyntaxError, Net::ProtocolError => e
         f_data[JSON_RESPONSE] = JSON_ERROR
         f_data[JSON_DATA][JSON_MESSAGE] = e.to_s
         if $options[:verbose]
            puts "#{e.to_s}\n\n#{e.backtrace}"
         end
      end
      
      return f_data
   end
   
   #
   # Dibuja un borde dentro de una hoja de Excel.
   #
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   # @param p_border_color [String] Color del borde en formato HTML
   # @param p_top_left [String] Celda del borde superior izquierdo
   # @param p_bottom_right [String] Celda del borde inferior derecho 
   # @example Borde azul con esquina superior izquierda en A5 y esquina inferior derecha en AB10
   #   f_player.draw_spreadsheet_border(p_workbook, p_worksheet, '#3498db', 'A5', 'AB10')
   # @author Gerard Carrasquer
   #
   def draw_spreadsheet_border(p_workbook, p_worksheet, p_border_color, p_top_left, p_bottom_right)
    
      f_top_left_letters = p_top_left.match(/[A-Za-z]+/).to_a.last.to_s
      f_bottom_right_letters = p_bottom_right.match(/[A-Za-z]+/).to_a.last.to_s
      f_top_left_number = p_top_left.match(/\d+/).to_a.last.to_i
      f_bottom_right_number = p_bottom_right.match(/\d+/).to_a.last.to_i
      
      
      
      #Border formats
      f_top_border_format = p_workbook.add_format(:top => 5, :border_color => p_border_color)
      f_top_left_border_format = p_workbook.add_format(:left => 5, :top => 5, :border_color => p_border_color)
      f_top_right_border_format = p_workbook.add_format(:right => 5, :top => 5, :border_color => p_border_color)
      f_bottom_border_format = p_workbook.add_format(:bottom => 5, :border_color => p_border_color)
      f_bottom_left_border_format = p_workbook.add_format(:left => 5, :bottom => 5, :border_color => p_border_color)
      f_bottom_right_border_format = p_workbook.add_format(:right => 5, :bottom => 5, :border_color => p_border_color)
      f_left_border_format = p_workbook.add_format(:left => 5, :border_color => p_border_color)
      f_right_border_format = p_workbook.add_format(:right => 5, :border_color => p_border_color)
      
      #Top and bottom borders
      f_letter = f_top_left_letters.clone
      while f_letter != f_bottom_right_letters
         
         p_worksheet.write("#{f_letter}#{f_top_left_number}", ' ', f_top_border_format)
         p_worksheet.write("#{f_letter}#{f_bottom_right_number}", ' ', f_bottom_border_format)
         
         f_letter.succ!
      end

      #Corners
      p_worksheet.write("#{f_bottom_right_letters}#{f_top_left_number}", ' ', f_top_right_border_format)
      p_worksheet.write(p_top_left, ' ', f_top_left_border_format)
      p_worksheet.write("#{f_top_left_letters}#{f_bottom_right_number}", ' ', f_bottom_left_border_format)
      p_worksheet.write(p_bottom_right, ' ', f_bottom_right_border_format)
      
      #Left and right borders
      (f_top_left_number+1).upto(f_bottom_right_number-1) do |n|
         p_worksheet.write("#{f_bottom_right_letters}#{n}", ' ', f_right_border_format)
         p_worksheet.write("#{f_top_left_letters}#{n}", ' ', f_left_border_format)
      end
      
 
      
   end
   
   #
   # Obtiene el número de la jornada actual.
   #
   # @return [Integer, nil] El número de la jornada actual o nil en caso de error.
   # @raise [Timeout::Error, Errno::EINVAL, Errno::ECONNRESET, EOFError,
   #          Net::HTTPBadResponse, Net::HTTPHeaderSyntaxError, Net::ProtocolError] Error en la descarga del 
   #          fichero JSON que contiene la jornada actual
   # @author Gerard Carrasquer
   #
   def get_current_week
      f_current_week_data = get_data("#{FANTASY_API_SERVER}/#{FANTASY_WEEK_CURRENT_URL}")
      f_current_week_number = nil
      
      if f_current_week_data[JSON_RESPONSE] == JSON_OK
         f_current_week_number = f_current_week_data[JSON_DATA][JSON_PREVIOUS_WEEK]
      end
      
      return f_current_week_number
   end
   
   #
   # Inserta un gráfico en una hoja de excel basándose en datos en columnas.
   # @param p_workbook [WriteXLSX] Objeto que representa la hoja de excel
   # @param p_worksheet [Writexlsx::Worksheet] Objeto que representa una pestaña de la hoja de excel
   # @param p_type [String] Tipo de gráfico ('area', 'bar', 'column', 'line', 'pie', 'doughnut'
   #  'scatter', 'stock', 'radar')
   # @param p_cell [String] Celda del borde superior izquierdo del gráfico
   # @param p_title [String] Título del gráfico
   # @param p_x_axis [String] Título del eje x
   # @param p_y_axis [String] Título del eje y
   # @param p_series_data [Array<(Hash)>] Series de datos del gráfico 
   # @option p_series_data [String] :name Nombre de la serie
   # @option p_series_data [String] :categories_column Columna que contiene las categorías de datos
   # @option p_series_data [String] :values_column Columna que contiene los datos
   # @option p_series_data [Integer]:init_line Línea inicial de los datos
   # @option p_series_data [Integer] :end_line Línea final de los datos
   # @example Gráfico en la celda A5 y dos series de datos en las columnas C y D entre las líneas 10..20 y 40..50  
   #   draw_column_chart(f_workbook, f_points_chart_worksheet, f_worksheet.name, 
   #                  'column', "A5", "Título", 'Eje x', 'Eje y', 
   #                  {:name => 'Serie', :categories_column => 'C', :values_column => D, :init_line => 10, :end_line => 20},
   #                  {:name => 'Puntos', :categories_column => 'C', :values_column => D, :init_line => 40, :end_line => 50})
   # @author Gerard Carrasquer
   #
   
   def draw_column_chart(p_workbook, p_worksheet, p_worksheet_name, p_type, p_cell, p_title, p_x_axis, p_y_axis, 
            *p_series_data)

      f_chart = p_workbook.add_chart(:type => p_type, :embedded => 1)
      f_chart.set_size(:x_scale => 1.5, :y_scale => 2)
      #f_chart.set_legend(:none => 1)
      
      p_series_data.each do |s|
         f_chart.add_series(
         :name => s[:name],
         :categories  => "=#{p_worksheet_name}!$#{s[:categories_column]}$#{s[:init_line]}:$#{s[:categories_column]}$#{s[:end_line]}",
         :values => "=#{p_worksheet_name}!$#{s[:values_column]}$#{s[:init_line]}:$#{s[:values_column]}$#{s[:end_line]}",
         :data_labels => { :value => 1 }
         )
         
      end
      
         
      f_chart.set_title(:name  => p_title)
      f_chart.set_x_axis(:name => p_x_axis)
      f_chart.set_y_axis(:name => p_y_axis)

      # Insert the chart into the worksheet (with an offset).
      p_worksheet.insert_chart(
         p_cell, f_chart,
         :x_offset => 5, :y_offset => 10
         )
   end
end
