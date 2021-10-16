require './fantasy_constants.rb'
require './fantasy.rb'
require './player.rb'

class Players < Fantasy

   def initialize
      @players_data = get_data("#{FANTASY_API_SERVER}/#{FANTASY_PLAYERS_URL}")
      
      if @players_data[JSON_RESPONSE] == JSON_ERROR
         puts "Error: #{@players_data[JSON_DATA][JSON_MESSAGE]}"
      else
         @players_data[JSON_DATA].each do |p|
            print "Descargando jugador id #{p[JSON_ID]}: #{p[JSON_NICKNAME]}.... "
            player = Player.new(p[JSON_ID])
            print "#{player[JSON_RESPONSE]}\n"
            
            if player[JSON_RESPONSE] == JSON_OK
               player.to_xlsx("#{$options[:players_folder]}/#{p[JSON_NICKNAME].gsub(/[^0-9A-Z]/i, '_')}.xlsx")
            end
         end
      end
   end
end
