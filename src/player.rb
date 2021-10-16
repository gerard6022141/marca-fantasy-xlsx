require './fantasy_constants.rb'
require './fantasy.rb'
require 'write_xlsx'
require 'down'

class Player < Fantasy

   #Constructor de la clase
   #Descarga la ficha del jugador y la guarda en un Hash
   
   def initialize(p_player_id)
      super()
      @player_data = Hash.new
      
      @player_data = get_data("#{FANTASY_API_SERVER}/#{FANTASY_PLAYER_URL}#{p_player_id}")
            
   end
   
   def [](p_json_tag)
      return @player_data[p_json_tag]
   end
   
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

         f_worksheet = f_workbook.add_worksheet
         
         player_head_to_xlsx(f_workbook, f_worksheet)
         
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
   
   def player_head_to_xlsx(p_workbook, p_worksheet)
      
      f_format = p_workbook.add_format
      
      f_format.set_bold
      f_format.set_color('#ec7063')
      f_format.set_font('Calibri')
      f_format.set_size(16)   
      f_format.set_top(5)
      f_format.set_border_color('#aeb6bf')
      
      p_worksheet.write('C2', ' ', f_format)
      p_worksheet.write('D2', ' ', f_format)
      p_worksheet.write('E2', ' ', f_format)  
      p_worksheet.write('F2', ' ', f_format)
      
      f_format.set_right(5)
      
      p_worksheet.write('G2', ' ', f_format)
      
      #f_format.set_top(0)
      
      p_worksheet.write('G3', '', f_format)
       
      
      f_tempfile = Down.download(@player_data[JSON_DATA][JSON_IMAGES][JSON_TRANSPARENT][JSON_128X128])
      p_worksheet.insert_image(2, 2, f_tempfile)
      
      p_worksheet.set_column(4, 4, 40)
      p_worksheet.set_row(3, 20)
      
      p_worksheet.write('E4', @player_data[JSON_DATA][JSON_NAME], f_format)
      
      
      
   end
end
