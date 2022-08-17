require 'roo'
require 'tiny_tds'
require_relative 'logging'
require 'YAML'
#----------------------------------------------------------------------------------
# CONNECTION STRING
client = TinyTds::Client.new username: '**', password: '*********',
                             host: '*********', port: 1433,
                             database: '**********', azure:false
#----------------------------------------------------------------------------------
# CONFIG PARAMETERS
cfg =YAML.load(File.read("config_LAP.yml"))
table = cfg[:table]
exc = cfg[:active] #Change to 0 to test in cfg
filenames = Dir.entries(cfg[:path])#location of files

#----------Truncs Table on run------------#
# if exc == 1
#   client.execute("TRUNCATE TABLE #{table}")
# end


def main_prog(filename, cfg, table, exc, client)
  bom = Roo::Spreadsheet.open(filename)
  bom.sheets

  ################Insert/Format Box######################


  def insert_sql(part, exc, table, sku, client, comm, sheet_name, col, part_start, prt_qty, bom, sheet_nbr, colm_nbr)
    ii = 0


    while ii < (sku.length)

      part_num =  "#{ part[col].to_s + part[col+1].to_s + part[col+2].to_s.gsub('*', '%') }" #####CHANGED HERE
      part_desc = part[col-2].to_s

      tmp_part = part[prt_qty..(part.length)]
      part_proc = tmp_part[ii]

      if (part_proc != nil) && (sku[ii] != nil)
        two_sku = sku[ii-1].to_s.split(/\n/)#Splits Sku into array  ####CHANGE HERE

        two_sku.each{ |splitsku|  #adds split Sku and submits
                                  #puts splitsku
        value_submission = " '#{part_num}', '#{part_desc}', '#{splitsku}', #{part_proc}, '#{sheet_name}'"
                                  #puts part[col]
        comm_submission = comm
                                  #-------Submit results to DB-------#
        if exc == 1

          result = client.execute("INSERT INTO #{table}(#{comm_submission}) VALUES (#{value_submission})")
          begin
            result.insert
            log = Logging.new
            message = "Success #{comm_submission} #{value_submission}"
            log.create_entry message, 1
          rescue => error
            log = Logging.new
            message = "Error #{comm_submission} #{value_submission}
            #{error.inspect}"
            log.create_entry message, 1
          end
        else
          log = Logging.new
          message = "Fail #{comm_submission} #{value_submission}"
          log.create_entry message, 1
          puts message
        end
        }
      end
      ii+=1
    end
  end

  ###############################################
  # CREATE THE TABLE IF IT DOESN'T EXIST
  inc = 0 #Change to 0 to test

  table_question = client.execute("
IF (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME= '#{table}'))
    begin
        SELECT '1'
        end
    ELSE SELECT '0'")

  table_exists =''

  table_question.each(:as => :array) do |x|
    table_exists =  x
  end



  if table_exists[0].to_s == '0'

    field_string = "id int PRIMARY KEY IDENTITY (1, 1), Part_Num VARCHAR(100) , Part_Desc VARCHAR(250), Item_Nbr VARCHAR(50), Qty VARCHAR(20), [Group] VARCHAR(100)"

    new_tab = client.execute("CREATE TABLE dbo.#{table}(#{field_string})")
    new_tab.do
  end


  #----------------------------------------------------------------------------------
  # GENERATE TABLE COLUMNS FOR INSERT STATEMENT
  def array_to_int(value)
    change = value.join("],[")
    change_2 = "["+change+"]"
  end

  result = client.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'#{table}' and COLUMN_NAME <> 'id'")
  storage = []
  result.each(:as => :array) do |x|

    storage << x
  end
  comm_submission  =  (array_to_int(storage) )



  ################Array Box######################
  #-------------Find Sheet Number-------------#
  iii = 0
  until iii == 1 #NO LONGER NEEDED

    sheet_nbr = 0
    colm_nbr = 0
    sheet_name = ''
    inc = 0
    until inc == (bom.sheets.length)

      if bom.sheets[inc].to_s.upcase.include? "22"
        bom.sheet(inc)


        sheet_name = bom.sheets[inc].to_s

      end

      inc+=1
    end

    #-------------Get Sku-------------#
    ii = 0
    until ii == 15
      row = bom.row(ii)
      colm = bom.column(ii)
      row.each{|it|
        if it.to_s.upcase.include? "SKU"

          sheet_nbr=ii
        end}
      colm.each { |it|
        #puts it
        if it.to_s.upcase.include? "SKU"

          colm_nbr = ii #CHANGE HERE
        end
      }

      ii+=1

    end

    sku = []
    crt_sku = bom.row(sheet_nbr)#.to_s
    crt_sku = crt_sku[(colm_nbr)..crt_sku.length]
    crt_sku.each{|it|
      if it != nil and (it.to_s.upcase.include? "SKU") == false
        sku<<it
        #puts it
      end}
    #-------part array selection-------#
    def get_part_arr(bom, i)
      part = bom.row(i)

    end

    ################Execution Box######################

    #---------FIND SPECIFIC CELLS---------#
    def find_cell_cont(phrase, bom)
      #takes search terms and looks for cell containing terms
      start= 0
      if phrase != nil
        i = 0
        cell = ('A'..'Z').to_a
        while i <= 27
          inc = 1
          while inc < 20
            if bom.cell(cell[i], inc).to_s.include? phrase

              start = [i, inc]
            end
            inc +=1
          end
          i+=1
        end
      end
      start
    end

    #----------Establish starting cells-----------#

    part_start = find_cell_cont('Part Description', bom)[0]#Term for search
    part_down = find_cell_cont('Part Description', bom)[1]

    col = find_cell_cont('Part Number', bom)[0]  #Find numeric index for column containing phrase
    prt_qty = find_cell_cont('SKU', bom)[0]


    #-----------------Sends excel data to be processed to DB-------------------#

    i=part_down#(part_start -1)
    while i <= bom.column(col).length
      part = get_part_arr(bom,  i) #Part is and array of a row

      unless part[1].to_s == nil
        #puts part[0..20].to_s
        if bom.cell(i, col+1) != nil

          insert_sql(part, exc, table, sku, client, comm_submission, sheet_name, col, part_start, prt_qty, bom, sheet_nbr, colm_nbr)
        end

      end

      i+=1

    end


    iii+=1
  end

end
############RUN ALL FILES##############
#----------Runs each file in folder------------#
filenames = filenames[5]
#filenames.each do |it|
file = ("#{cfg[:path].to_s}#{filenames}")
main_prog(file, cfg, table, exc, client)
#end
