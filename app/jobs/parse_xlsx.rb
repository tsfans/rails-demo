#!/usr/bin/ruby

require 'roo'

class ParseXlsx

  @@category=['OPEX', 'VM', 'PMC']
  @@room_amenities_row_numbers={1 => 19, 2 => 15, 3 => 96, 4 => 103, 5 => 104, 7 => 97, 8 => 100, 12 => 18, 13 => 46,
                                14 => 24, 15 => 22, 16 => 20, 17 => 24}
  @@room_info_attributes_row_numbers={1 => 10, 2 => 11, 3 => 12, 4 => 24, 5 => 20, 6 => 22, 7 => 63, 8 => 63, 9 => 83,
                                      10 => 59, 11 => 60, 12 => 61, 13 => 108, 14 => 91, 15 => 107, 16 => 64, 17 => 98}
  @@room_amenities_row_numbers_china = {3 => 61, 8 => 65, 5 => 69}
  @@room_info_attributes_row_numbers_china = {17 => 63, 8 => 41, 12 => 39, 10 => 37, 11 => 38, 1 => 5, 2=> 6, 3 => 7}
  @@hotel_amenities_row_numbers={1 => 73, 2 => 75, 3 => 76, 4 => 77, 5 => 78, 6 => 79, 7 => 80, 8 => 81, 9 => 65, 11 => 55, 12 => 36}
  @@hotel_info_attributes_row_numbers={1 => 72, 2 => 29, 3 => 66, 4 => 68, 5 => 69}
  @@room_number_row=8
  @@room_number_row_china=3
  @@room_category_row_china=10
  @@floor_number_row_china=4
  @@max_occupancy_row=108
  @@ac_exceptions=['non-ac-room']
  @@washroom_type = {'西式' => 'Western', '英印式' => 'Anglo-Indian', '印度式' => 'Indian', '独立式' => 'Non-Attached', 'Western' => 'Western', 'Anglo-Indian' => 'Anglo-Indian', 'Indian' => 'Indian', 'Non-Attached' => 'Non-Attached'}



        def parse_file
            file = Roo::Spreadsheet.open('/Users/oyo/Documents/audit_sheet.xlsx')
            room_sheet = file.sheets.map {|sheet_name| sheet_name.downcase.include?('room') ? sheet_name : nil}.compact.last
            sheets = file.sheets.map {|i| @@category.include?(i.split('_').last) ? i : nil}.compact
            room_sheet = file.sheets.map {|sheet_name| sheet_name.downcase.include?('room') ? sheet_name : nil}.compact.last
            parse_room_amenities_and_info_china(file, room_sheet)

        end
        
        def parse_room_amenities_and_info_china(file, room_sheet)
            file.default_sheet=room_sheet
            floor_numbers=parse_numbers(file, @@floor_number_row_china)
            room_numbers=parse_numbers(file, @@room_number_row_china)
            room_category_row=file.row(@@room_category_row_china)
            room_catrgories=room_category_row[2..-1].compact
            
            # rooms=hotel_audit.hotel_audit_rooms.includes(:hotel_audit_room_amenities, :hotel_audit_rooms_info_params).
            #  index_by(&:room_number)

            puts floor_numbers.size
            puts room_numbers.size
            puts room_catrgories.size


            hotel_audit_rooms_attributes=Array.new
            (0..room_numbers.size-1).each do |i|
              hotel_audit_rooms_attributes.push(
               {
                floor_number: floor_numbers[i], 
                room_number: room_numbers[i], 
                audit_status: "Audited",
                byg_status: "Green",
                room_category: room_catrgories[i],
                view: "None",
                hotel_audit_rooms_info_params_attributes: [],
                hotel_audit_room_amenities_attributes: []
              })
             end 
            # puts hotel_audit_rooms_attributes
            @@room_amenities_row_numbers_china.each do |amenity_id, row_number|
              amenity_values=file.row(row_number)[3..-1][0..room_numbers.length-1]
              (0..amenity_values.length-1).each do |i|
                begin
                  # audit_amenity=rooms[(hotel_audit_rooms_attributes[i][:room_number]).to_s].hotel_audit_room_amenities.
                  # select {|amenity| amenity[:room_amenity_id] == amenity_id}.last
                  # id=audit_amenity.nil? ? nil : audit_amenity.id
                  if [1, 2, 12, 13].include? amenity_id
                    hotel_audit_rooms_attributes[i][:hotel_audit_room_amenities_attributes].
                      push(handle_amenities_values(amenity_id, amenity_values[i], audit_amenity.try(:hotel_audit_room_amenities_values)))
                  elsif [15, 16].include? amenity_id
                    is_available=amenity_values[i].to_i>0 ? true : false
                    hotel_audit_rooms_attributes[i][:hotel_audit_room_amenities_attributes].
                      push({'room_amenity_id' => amenity_id, 'is_available' => is_available})
                  elsif amenity_id==14
                    is_available=amenity_values[i].to_i==1 ? true : false
                    hotel_audit_rooms_attributes[i][:hotel_audit_room_amenities_attributes].
                      push({'room_amenity_id' => amenity_id, 'is_available' => is_available})
                  elsif amenity_id==17
                    is_available=amenity_values[i].to_i>1 ? true : false
                    hotel_audit_rooms_attributes[i][:hotel_audit_room_amenities_attributes].
                      push({'room_amenity_id' => amenity_id, 'is_available' => is_available})
                  else
                    is_available=(amenity_values[i].to_s == '是')? true : false
                    hotel_audit_rooms_attributes[i][:hotel_audit_room_amenities_attributes].
                    push({'room_amenity_id' => amenity_id, 'is_available' => is_available})
                  end
              rescue
                raise "Error occurred in row #{row_number} in #{room_sheet}"
              end
            end
          end

          @@room_info_attributes_row_numbers_china.each do |attribute_id, row_number|
            attribute_values=file.row(row_number)[3..-1][0..room_numbers.length-1]
            # if attribute_values.present?
              (0..attribute_values.length-1).each do |i|
                begin
                  # audit_attribute=rooms[(hotel_audit_rooms_attributes[i][:room_number]).to_s].hotel_audit_rooms_info_params.
                  #   select {|attribute| attribute[:room_info_params_attribute_id] == attribute_id}.last
                  # id=audit_attribute.nil? ? nil : audit_attribute.id
                  if [14, 17].include? attribute_id
                    value=((attribute_values[i].to_s=='y' || attribute_values[i] == '是') ? 'Yes' : 'No')
                  elsif attribute_id==16
                    value=((attribute_values[i].to_s=='y' || attribute_values[i] == '是') ? 'No' : 'Yes')
                  elsif attribute_id==7
                    value= (attribute_values[i]=='Non-Attached' ? 'No' : 'Yes')
                  elsif attribute_id==8
                    value = (@@washroom_type[attribute_values[i].to_sym] =='Non-Attached' ? nil : @@washroom_type[attribute_values[i].to_sym])
                  elsif attribute_id==9
                    value = ((attribute_values[i]==nil || attribute_values[i].to_i<=0) ? 'No' : 'Yes')
                  elsif attribute_id==15
                    value=attribute_values[i].to_s
                    value='other' if value==nil || value=='none'
                  else
                    if [1, 2, 3, 13].include?(attribute_id) && attribute_values[i].to_s.strip==nil
                      raise "Something is missing in row #{row_number} in #{room_sheet}"
                    end
                    value= ([1, 2, 3, 10, 11, 12, 15].include?(attribute_id) ? attribute_values[i].to_s : attribute_values[i].to_i)
                  end
                  hotel_audit_rooms_attributes[i][:hotel_audit_rooms_info_params_attributes].
                  push({'room_info_params_attribute_id' => attribute_id, 'value' => value})
                rescue
                  raise "Error occurred in row #{row_number} in #{room_sheet}"
                end
              end
            # end
          end

          puts hotel_audit_rooms_attributes
        end

        def handle_amenities_values(amenity_id, value, values)
          amenity_values_attributes=[]
          if [12, 13].include? amenity_id
            is_available=((value.blank? || value.to_s.try(:downcase)=='n') ? true : false)
          elsif amenity_id==2
            is_available=(@@ac_exceptions.include?(value.try(:lowercase)) ? false : true)
          elsif amenity_id==1
            type='LED/LCD'
            size=(value.to_i>21 ? value.to_i : 24)
            is_available=size.present?
            if values.present?
              size_value_id=values.where(:room_amenities_key_id => 1).last.try(:id)
              type_value_id=values.where(:room_amenities_key_id => 2).last.try(:id)
            else
              size_value_id=nil
              type_value_id=nil
            end
            amenity_values_attributes=[{:room_amenities_key_id => 1, :value => size,:id=>size_value_id},
                                       {:room_amenities_key_id => 2, :value => type,:id=>type_value_id}]
          else
            if values.present?
              no_seats_id=values.where(:room_amenities_key_id => 3).last.try(:id)
            else
              no_seats_id=nil
            end
            amenity_values_attributes=[{:room_amenities_key_id => 3, :value => value.to_i,:id=>no_seats_id}]
            is_available=(value.present? && value != 0) ? true : false
          end
          return {'room_amenity_id' => amenity_id, 'is_available' => is_available,
                  'hotel_audit_room_amenities_values_attributes'=>amenity_values_attributes}
    end

        def parse_numbers(file, row_number)
          number_rows=file.row(row_number)
          numbers=number_rows[2..-1].compact
          if numbers.first.is_a? Float
            numbers=numbers.map {|i| i.to_i.to_s}
          else
            numbers=numbers.map {|i| i.to_s}
          end
          numbers.delete('0')
          numbers.delete('0.0')
          return numbers
        end

        def array_compact
          arr1=Array[1,2,3,4,5]
          arr2=Array[2,5]
          arr1=arr1.select{|n| !arr2.include?(n)}
          puts(arr1)
          puts("123456"[/\d+/])

        end

end


object = ParseXlsx.new
# object.parse_file
object.array_compact








