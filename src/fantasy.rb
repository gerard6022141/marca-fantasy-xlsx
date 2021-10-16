require 'uri'
require 'net/http'
require 'json'
require './fantasy_constants.rb'

class Fantasy

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
      rescue => e
         f_data[JSON_RESPONSE] = JSON_ERROR
         f_data[JSON_DATA][JSON_MESSAGE] = e.to_s
         if $options[:verbose]
            puts "#{e.to_s}\n\n#{e.backtrace}"
         end
      end
      
      return f_data
   end
   
   
end
