require 'optparse'
require './players.rb'

$options = Hash.new
$parse_ok=true

optparse = OptionParser.new do |opts|
                
   opts.banner = "Uso: ruby marca-fantasy-ods.rb [options]"
      
      opts.separator "Opciones:"
        
        
      opts.on( '-v', '--verbose', 'Muestra más información en pantalla' ) do
             $options[:verbose]=true
      end     
     
     opts.on( '-f', '--players_folder FILE', 'Carpeta de ficheros de jugadores' ) do |folder|
             $options[:players_folder]=folder
     end
        
     opts.on( '-h', '--help', 'Muestra esta pantalla' ) do     
             puts opts
             $parse_ok=false
     end
end

begin
   optparse.parse!(ARGV)
   if $options[:players_folder].nil?
      raise "Falta parámetro: --players_folder"
      $parse_ok=false
   elsif not Dir.exist?($options[:players_folder])
      raise "El directorio #{$options[:players_folder]} no existe"
      $parse_ok=false
   end
        
rescue OptionParser::MissingArgument=>e
        puts e
        puts optparse
        $parse_ok=false
rescue =>e
        puts e
        #puts optparse
        $parse_ok=false
end

if $parse_ok
   players = Players.new
   
end
