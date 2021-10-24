require 'optparse'
require './players.rb'

#
# Programa principal
#

$options = Hash.new
$parse_ok=true

optparse = OptionParser.new do |opts|
                
   opts.banner = "Uso: ruby marca-fantasy-xlsx.rb [options]"
      
    opts.separator "Opciones:"
        
        
    opts.on( '-v', '--verbose', 'Muestra más información en pantalla (opcional)' ) do
        $options[:verbose]=true
    end     
     
    opts.on( '-f', '--players_file FILE', 'Genera un fichero que contiene todos los jugadores' ) do |file|
        $options[:players_file]=file
    end
     
    opts.on( '-p', '--players_folder FOLDER', 'Genera un fichero por cada jugador en la carpeta especificada' ) do |folder|
        $options[:players_folder]=folder
    end
     
    opts.on( '-i', '--players_id LIST',Array, 'Filtra la descarga los identificadores de jugadores especificados' ) do |params|
        $options[:players_id] = params
    end
    
    opts.on( '-s', '--search_names LIST',Array, 'Filtra la descarga a los nombres de jugadores especificados' ) do |params|
        $options[:players_names] = params
    end
     
     opts.on( '-c','--chart', 'Genera una pestaña de gráficas' ) do
        $options[:chart]=true
    end 
    
    opts.on( '-m', '--compare_players', 'Compara jugadores (opcional)' ) do
        $options[:compare_players]=true
    end
    
    opts.on( '-w', '--weeks LISTA',Array, 'Filtra las jornadas especificadas para las estadísticas' ) do |params|
        $options[:weeks] = Array.new
        
        if !params.nil?
          params.each do |p|
            if p =~ /(\d+)\.\.(\d+)/
              $~[1].to_i.upto($~[2].to_i) do |i|
                $options[:weeks].push(i)
              end
              
            elsif p=~ /\d+/
              $options[:weeks].push($~[0].to_i)
            end
          end
        end
    end 
    
        
     opts.on( '-h', '--help', 'Muestra esta pantalla' ) do     
             puts opts
             $parse_ok=false
     end
end

begin
   optparse.parse!(ARGV)
   if $options[:players_folder].nil? && $options[:players_file].nil?
      raise "Falta parámetro: --players_folder o --players_file"
      $parse_ok=false
   elsif $options[:players_file].nil? && $options[:compare_players]
      raise "El parámetro --players_file es obligatorio cuando se usa --compare_players"
      $parse_ok=false
   elsif !$options[:weeks].nil? && $options[:weeks].empty?
      raise "El rango de jornadas especificado en --weeks es incorrecto"
      $parse_ok=false
   elsif !$options[:players_folder].nil? && !Dir.exist?($options[:players_folder])
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
