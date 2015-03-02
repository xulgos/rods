# coding: UTF-8
#
# = RODS - Ruby Open Document Spreadsheet
# This class provides a convenient interface for fast reading and writing 
# spreadsheets conforming to Open Document Format v1.1..
# Installiation of an office-application (LibreOffice, OpenOffice.org) is not required as the code directly 
# manipulates the XML-files in the zipped *.ods-container. 
#
# = Copyright
# Copyright (c) <em>Dr. Heinz Breinlinger</em> (2011).
# Licensed under the same terms as Ruby. No warranty is provided.
#
# = Tutorial
# Please refer to README for how to use the interface with many annotated examples.
# 
require 'rubygems'
require 'zip/zipfilesystem'
require 'rexml/document'
require 'helpers'
require 'color'
require 'exceptions'

module Rods
  ROW = "row"
  CELL = "cell"
  COLUMN = "column"
  TAG = "tag"
  TEXT = "text"
  CHILD = "child"
  STYLES = "styles"
  CONTENT = "content"
  DUMMY = "dummy"
  WIDTH = "width"
  NODE = "node"
  BEFORE = "before"
  AFTER = "after"
  INDEX = "index"
  NUMBER = "number"
  BOTH = "both"
  WIDTHEXCEEDED = "exceeded"

  class Document
    ##########################################################################
    # Convenience-function to switch the default-style for the display of
    # date-values. The switch is valid for all subsequently created cells with 
    # date-values.  
    # Builtin valid values are
    # * 'date_style'
    #   * -> "02.01.2011" (German formatting)
    # * 'date_day_style' 
    #   * -> "Su"
    # Example
    #   sheet.set_date_format "date_day_style"  # RODS' default format for display of weekday
    #   sheet.set_date_format "date_style"     # RODS' default format for date  "12.01.2011" German format
    #-------------------------------------------------------------------------
    def set_date_format format_name
      case format_name
        when "date_style" then @date_style = "date_style"
        when "date_day_style" then @date_style = "date_day_style"
        else die "invalid format-name #{format}"
      end
    end
    ##########################################################################
    # internal: Error-method for displaying fatal error-message
    #-------------------------------------------------------------------------
    def die message
      raise RodsError, message
    end
    ##########################################################################
    # internal: Returns a new REXML::Element of type 'cell' with repetition-attribute set to 'n'
    #-------------------------------------------------------------------------
    def create_cell repetition
      return create_element CELL,repetition
    end
    ##########################################################################
    # internal: Returns a new REXML::Element of type 'row' with repetition-attribute set to 'n'
    #-------------------------------------------------------------------------
    def create_row repetition
      return create_element ROW,repetition
    end
    ##########################################################################
    # internal: Returns a new REXML::Element of type 'column' with repetition-attribute set to 'n'
    #-------------------------------------------------------------------------
    def create_column repetition
      return create_element COLUMN,repetition
    end
    ##########################################################################
    # internal: Returns a new REXML::Element of type 'row', 'cell' or 'column'
    # with repetition-attribute set to 'n'
    #-------------------------------------------------------------------------
    def create_element type, repetition
      if repetition < 1
        die "invalid value for repetition #{repetition}"
      end
      if type == ROW
        row = REXML::Element.new "table:table-row"
        if repetition > 1
          row.attributes["table:number-rows-repeated"] = repetition.to_s
        end
        return row
      elsif type == CELL
        cell = REXML::Element.new "table:table-cell"
        if repetition > 1
          cell.attributes["table:number-columns-repeated"] = repetition.to_s
        end
        return cell
        elsif type == COLUMN
        column = REXML::Element.new "table:table-column"
        if repetition > 1
          column.attributes["table:number-columns-repeated"] = repetition.to_s
        end
        column.attributes["table:default-cell-style-name"] = "Default"
        return column
      #----------------------------------------------
      else
        die "Invalid Type: #{type}"
      end
    end
    ##########################################################################
    # internal: Sets repeption-attribute of REXML::Element of type 'row' or 'cell' 
    #------------------------------------------------------------------------
    def set_repetition element, type, repetition
      die "wrong type #{type}" if type != ROW && type != CELL
      die "invalid value for repetition #{repetition}" if repetition < 1
      die "element is nil" unless element
      kind_of_repetition = type == ROW ?  "table:number-rows-repeated" : "table:number-columns-repeated"
      if repetition.to_i == 1
        element.attributes.delete kind_of_repetition
      else
        element.attributes[kind_of_repetition] = repetition.to_s
      end
    end
    ##########################################################################
    # Writes the given text to the cell with the given indices.
    # Creates the cell if not existing.
    # Formats the cell according to type and returns the cell.
    #   cell = sheet.write_get_cell 3,3,"formula:time"," = C2-C1"
    # This is useful for a subsequent call to 
    #   sheet.set_attributes cell, { "background-color" => "yellow3"}
    #-------------------------------------------------------------------------
    def write_get_cell row_ind, col_ind, type, text
      cell = get_cell rowInd, colInd
      write_text cell, type, text
      return cell
    end
    ##########################################################################
    # Writes the given text to the cell with the given indices.
    # Creates the cell if not existing.
    # Formats the cell according to type.
    #   sheet.write_cell 1,1,"date","31.12.2010" # 1st row, 1st column
    #   sheet.write_cell 2,1,"formula:date"," = A1+1"
    #   sheet.write_cell 1,3,"time","13:37" # German time-format
    #   sheet.write_cell 1,4,"currency","19,99" # you could also use '.' as a decimal separator
    #-------------------------------------------------------------------------
    def write_cell row, col, type, text
      cell = get_cell row, col
      write_text cell, type, text
    end
    ##########################################################################
    # Writes the given text to the cell with the given index in the given row.
    # Row is a REXML::Element.
    # Creates the cell if not existing.
    # Formats the cell according to type and returns the cell.
    #   row = sheet.get_row(17)
    #   cell = sheet.writeGetCellFromRow(row,4,"formula:currency"," = B5*1,19")
    #-------------------------------------------------------------------------
    def writeGetCellFromRow(row,colInd,type,text)
      cell = getCellFromRow(row,colInd)
      write_text(cell,type,text)
      return cell
    end
    ##########################################################################
    # Writes the given text to the cell with the given index in the given row.
    # Row is a REXML::Element.
    # Creates the cell if it does not exist.
    # Formats the cell according to type.
    #   row = sheet.get_row(3)
    #   sheet.writeCellFromRow(row,1,"date","28.12.2010")
    #   sheet.writeCellFromRow(row,2,"formula:date"," = A1+3")
    #-------------------------------------------------------------------------
    def writeCellFromRow(row,colInd,type,text)
      cell = getCellFromRow(row,colInd)
      write_text(cell,type,text)
    end
    ##########################################################################
    # Returns the cell at the given index in the given row.
    # Cell and row are REXML::Elements.
    # The cell is created if it does not exist.
    #   row = sheet.get_row(15)
    #   cell = sheet.getCellFromRow(row,17) # 17th cell of 15th row
    # Looks a bit strange compared to
    #   cell = sheet.get_cell(15,17)
    # but is considerably faster if you are operating on several cells of the
    # same row as after locating the first cell of the row the XML-Parser can start 
    # from the node of the already found row instead of having to locate the
    # row over and over again.
    #-------------------------------------------------------------------------
    def getCellFromRow(row,colInd)
      return get_child_by_index(row,CELL,colInd)
    end
    ##########################################################################
    # Returns the cell at the given indices.
    # Cell is a REXML::Element.
    # The cell is created if it does not exist.
    #   cell = sheet.get_cell(14,37)
    #-------------------------------------------------------------------------
    def get_cell row_index, col_index
      row = get_row row_index
      get_child_by_index row, CELL, col_index
    end
    ##########################################################################
    # Returns the row at the given index.
    # Row is a REXML::Element.
    # The row is created if it does not exist.
    #      row = get_row 1
    #      1.upto 500 do |i|
    #        row = get_row i 
    #        text1,type1 = read_cell_from_row row,3  
    #        text2,type2 = read_cell_from_row row,4 # XML-Parser can start from row-node instead of root-node
    #        puts "Read #{text1} of #{type1} and #{text2} of #{type2}
    #      end
    #-------------------------------------------------------------------------
    def get_row row_index
      current_table = @tables[@current_table_name][NODE]
      get_child_by_index current_table, ROW, row_index
    end
    ##########################################################################
    # internal: returns the child REXML::Element of the given type 
    # ('row', 'cell' or 'column') and index within the parent-element.
    # The child is created if it does not exist.
    #------------------------------------------------------------------------
    def get_child_by_index parent, type, index
      i = 0
      last_element = nil
      if type != ROW && type != CELL && type != COLUMN
        die "wrong type #{type}"
      end
      if index < 1
        die "invalid index #{index}"
      end
      die "parent-element does not exist" unless parent
      if type == ROW
        kind_of_element = "table:table-row"
        kind_of_repetition = "table:number-rows-repeated"
      elsif type == CELL || type == COLUMN
        if index > @tables[@current_table_name][WIDTH]
          @tables[@current_table_name][WIDTH] = index
          @tables[@current_table_name][WIDTHEXCEEDED] = true
        end
        kind_of_repetition = "table:number-columns-repeated"
        case type
          when CELL then kind_of_element = "table:table-cell"
          when COLUMN then kind_of_element = "table:table-column"
          else die "internal error: when-clause-failure for type #{type}"
        end
      else
        die "wrong type #{type}"
      end
      parent.elements.each(kind_of_element) do |element|
        i += 1
        last_element = element
        if i == index
          if repetition = element.attributes[kind_of_repetition]
            num_empty_elements_after = repetition.to_i - 1
            if num_empty_elements_after < 1
              die "new repetition < 1"
            end
            set_repetition element, type, 1
            element.next_sibling = create_element type, num_empty_elements_after
          end
          return element 
        elsif i < index
          if repetition = element.attributes[kind_of_repetition]
            index_of_last_empty_element = i + repetition.to_i - 1
            if index_of_last_empty_element < index
              i = index_of_last_empty_element
            else
              num_empty_elements_before = index - i
              num_empty_elements_after = index_of_last_empty_element - index
              set_repetition element, type, num_empty_elements_before
              element.next_sibling = create_element type, 1
              if num_empty_elements_after > 0
                element.next_sibling.next_sibling = create_element type, num_empty_elements_after
              end
              return element.next_sibling
            end
          end
        end
      end
      num_empty_elements_before = index - i - 1
      if i > 0
        element = create_element type, 1
        if num_empty_elements_before > 0
          last_element.next_sibling = create_element type, num_empty_elements_before
          last_element.next_sibling.next_sibling = element
        else
          last_element.next_sibling = element
        end
        return element
      else
        if index == 1
          newElement = create_element type, 1
          parent.add newElement
          return newElement
        else
          newElement = create_element type, num_empty_elements_before
          parent.add newElement
          newElement.next_sibling = create_element type, 1
          return newElement.next_sibling
        end
      end
    end
    ##########################################################################
    # internal: Determines the number of tables, initializes the internal
    # table-administration via Hashes and sets the current default-table for
    # all subsequent operations (first table of spreadsheet).
    #-------------------------------------------------------------------------
    def init_house_keeping
      @spread_sheet = @content_text.elements["/office:document-content/office:body/office:spreadsheet"]
      die "Could not extract office:spreadsheet" unless @spread_sheet
      @num_tables = 0
      @spread_sheet.elements.each "table:table" do |table|
        table_name = table.attributes["table:name"]
        die "Could not extract table_name" if table_name.empty?
        @tables[table_name] = Hash.new
        @tables[table_name][NODE] = table
        @tables[table_name][WIDTH] = get_table_width table
        @tables[table_name][WIDTHEXCEEDED] = false
        @num_tables += 1
      end
      if @num_tables == 0
        insert_table "Table 1"
      end
      first_table = @spread_sheet.elements["table:table[1]"]
      @current_table_name = first_table.attributes["table:name"]
    end
    ##########################################################################
    # returns the list of table names
    # ------------------------------------------------------------------------
    def tableNames
      @tables.keys
    end
    ##########################################################################
    # Renames the table of the given name and updates the internal table-administration.
    #   sheet.rename_table "Table1","not needed"
    #-------------------------------------------------------------------------
    def rename_table old_name, new_name
      die "table '#{old_name}' does not exist" unless @tables.has_key? old_name
      node = @tables[old_name][NODE]
      node.attributes["table:name"] = new_name
      @tables[new_name] = @tables[old_name]
      @tables.delete old_name
      @current_table_name = new_name if old_name == @current_table_name
    end

    def table_count
      @tables.length
    end

    def current_table
      @current_table_name
    end
    ##########################################################################
    # Sets the table of the given name as the default-table for all subsequent
    # operations.
    #   sheet.set_current_table "example"
    #-------------------------------------------------------------------------
    def set_current_table table_name
      die "table '#{table_name}' does not exist" unless @tables.has_key? table_name
      @current_table_name = table_name
    end
    ##########################################################################
    # Inserts a table of the given name before the given spreadsheet and updates
    # the internal table-administration.
    #   sheet.insertTableBefore("table2","table1") 
    #-------------------------------------------------------------------------
    def insertTableBefore(relativeTableName,tableName)
      insertTableBeforeAfter(relativeTableName,tableName,BEFORE)
    end
    ##########################################################################
    # Inserts a table of the given name after the given spreadsheet and updates
    # the internal table-administration.
    #   sheet.insertTableAfter("table1","table2") 
    #-------------------------------------------------------------------------
    def insertTableAfter(relativeTableName,tableName)
      insertTableBeforeAfter(relativeTableName,tableName,AFTER)
    end
    ##########################################################################
    # internal: Inserts a table of the given name before or after the given spreadsheet and updates
    # the internal table-administration. The default position is 'after'.
    #   sheet.insertTableBeforeAfter("table1","table2",BEFORE) 
    #-------------------------------------------------------------------------
    def insertTableBeforeAfter(relativeTableName,tableName,position = AFTER)
      die("insertTableAfter: table '#{relativeTableName}' does not exist") unless (@tables.has_key?(relativeTableName))
      die("insertTableAfter: table '#{tableName}' already exists") if (@tables.has_key?(tableName))
      #-----------------------------------------
      # alte Tabelle ermitteln
      #-----------------------------------------
      @spread_sheet.elements["table:table"].each{ |element|
        puts("Name: #{element.attributes['table:name']}")
      }
      relativeTable = @spread_sheet.elements["*[@table:name = '#{relativeTableName}']"]
      die("insertTableAfter: internal error: Could not locate existing table #{relativeTableName}") unless (relativeTable) 
      #-----------------------------------------
      # Neues Tabellenelement zunaecht per se (i.e. unverankert)  erschaffen
      #-----------------------------------------
      newTable = REXML::Element.new("table:table")
      newTable.add_attributes({"table:name" =>  tableName,
                               "table:print" => "false",
                               "table:style-name" => "table_style"})
      #-----------------------------------------
      # Unterelemente anlegen und neue Tabelle
      # hinter vorherige einfuegen
      #-----------------------------------------
      write_xml(newTable,{TAG => "table:table-column",
                         "table:style" => "column_style",
                         "table:default-cell-style-name" => "Default"})
      write_xml(newTable,{TAG => "table:table-row",
                         "table:style-name" => "row_style",
                         CHILD => {TAG => "table:table-cell"}})
      case position
        when BEFORE then @spread_sheet.insert_before(relativeTable,newTable)
        when AFTER then @spread_sheet.insert_after(relativeTable,newTable)
        else die("insertTableBeforeAfter: invalid parameter #{position}")
      end
      #---------------------------------------------------------------------------
      # Tabellen-Hash aktualisieren
      #---------------------------------------------------------------------------
      @tables[tableName] = Hash.new()
      @tables[tableName][NODE] = newTable
      @tables[tableName][WIDTH] = get_table_width(newTable)
      @tables[tableName][WIDTHEXCEEDED] = false
      @num_tables += 1
    end
    ##########################################################################
    # Inserts a table of the given name at the end of the spreadsheet and updates
    # the internal table-administration.
    #   sheet.insert_table "example"
    #-------------------------------------------------------------------------
    def insert_table table_name
      die "table '#{table_name}' already exists" if @tables.has_key? table_name
      new_table = write_xml @spread_sheet, {
        TAG => "table:table",
        "table:name" => table_name,
        "table:print" => "false",
        "table:style-name" => "table_style",
        "child1" => { TAG => "table:table-column",
                      "table:style" => "column_style",
                      "table:default-cell-style-name" => "Default" },
        "child2" => { TAG => "table:table-row",
                      "table:style-name" => "row_style",
                      "child3" => { TAG => "table:table-cell" }}}
      @tables[table_name] = Hash.new
      @tables[table_name][NODE] = new_table
      @tables[table_name][WIDTH] = get_table_width new_table
      @tables[table_name][WIDTHEXCEEDED] = false
      @num_tables += 1
    end
    ##########################################################################
    # Deletes the table of the given name and updates the internal 
    # table-administration.
    #   sheet.delete_table "Table2"
    #-------------------------------------------------------------------------
    def delete_table table_name
      die "table '#{table_name}' cannot be deleted as it is the current table" if table_name == @current_table_name
      die "invalid table-name/not existing table: '#{table_name}'" unless @tables.has_key? table_name
      node = @tables[table_name][NODE]
      @spread_sheet.elements.delete node
      @tables.delete table_name
      @num_tables -= 1
    end
    ##########################################################################
    # internal: Calculates the current width of the current table.
    #-------------------------------------------------------------------------
    def get_table_width table
      die "current table does not contain table:table-column" unless table.elements["table:table-column"]
      table_name = table.attributes["table:name"]
      die "Could not extract table_name" if table_name.empty?
      num_columns_of_table = 0
      table.elements.each "table:table-column" do |table_column|
        num_columns_of_table += 1
        num_repetitions = table_column.attributes["table:number-columns-repeated"]
        if num_repetitions
          num_columns_of_table += num_repetitions.to_i - 1
        end
      end
      num_columns_of_table
    end
    ##########################################################################
    # internal: Adapts the number of columns in the headers of all tables 
    # according to the right-most valid column. This method is called when
    # the spreadsheet is saved.
    #------------------------------------------------------------------------
    def pad_tables
      @tables.each do |tableName, tableHash|
        table = table_hash[NODE]
        width = table_hash[WIDTH]
        num_columns_of_table = get_table_width table
        if table_hash[WIDTHEXCEEDED]
          die "current table does not contain table:table-column" unless table.elements["table:table-column"]
          last_table_column = table.elements["table:table-column[last ]"]
          if last_table_column.attributes["table:number-columns-repeated"]
            num_repetitions = last_table_column.attributes["table:number-columns-repeated"].to_i + width - num_columns_of_table
          else
            num_repetitions = width - num_columns_of_table + 1 # +1 as column itself count as repeat
          end
          last_table_column.attributes["table:number-columns-repeated"] = num_repetitions.to_s
          table_hash[WIDTHEXCEEDED] = false
        end
      end
    end
    ##########################################################################
    # internal: Verifies the format of a given time-string and converts it into
    # a proper internal representation.
    #-------------------------------------------------------------------------
    def time_to_time_val text
      unless text.match /^\d{2}:\d{2}(:\d{2})?$/
        die "wrong time-format '#{text}' -> expected: 'hh:mm' or 'hh:mm:ss'"
      end
      unless text.match /^[0-1][0-9]:[0-5][0-9](:[0-5][0-9])?$|^[2][0-3]:[0-5][0-9](:[0-5][0-9])?$/
        die "time '#{text}' not in valid range"
      end
      time = text.match /(\d{2}):(\d{2})(:(\d{2}))?/
      hour = time[1]
      minute = time[2]
      seconds = time[4]
      "PT#{hour}H#{minute}M#{seconds.nil? ? "00" : seconds}S"
    end
    ##########################################################################
    # internal: Divides by 100 and returns a string
    #----------------------------------------------------------------------
    def percent_to_percent_val text
      (text.to_f/100.0).to_s
    end
    ##########################################################################
    # internal: Converts a date-string of the form '01.01.2010' into the internal
    # representation '2010-01-01'.
    #----------------------------------------------------------------------
    def date_to_date_val text
      return text if text =~ /(^\d{4})-(\d{2})-(\d{2})$/
      die "Date #{text} does not comply with format dd.mm.yyyy" unless text.match /^\d{2}\.\d{2}\.\d{4}$/
      text =~ /(^\d{2})\.(\d{2})\.(\d{4})$/
      $3+"-"+$2+"-"+$1 
    end
    ##########################################################################
    # Returns the content and type of the cell at the index in the given row
    # as strings. Row is a REXML::Element.
    # If the cell does not exist, nil is returned for text and type.
    # Type is one of the following office:value-types
    # * string, float, currency, time, date, percent, formula
    # The content of a formula is it's last calculated result or 0 in case of a
    # newly created cell. The text is internally cleaned from currency-symbols and
    # converted to a valid (English) float representation (but remains a string)
    # in case of type "currency" or "float".
    #   amount = 0.0
    #   5.upto(8){ |i|
    #     row = sheet.get_row(i)
    #     text,type = sheet.readCellFromRow(row,i)
    #     sheet.writeCellFromRow(row,9,type,(-1.0*text.to_f).to_s)
    #     if(type == "currency")
    #       amount += text.to_f
    #     end
    #   }
    #   puts("Earned #{amount} bucks")
    #---------------------------------------------------------------
    def readCellFromRow(row,colInd)
      j = 0
      #------------------------------------------------------------------
      # Fuer alle Spalten
      #------------------------------------------------------------------
      row.elements.each("table:table-cell"){ |cell|
        j = j+1
        #-------------------------------------------
        # Spaltenwiederholungen addieren
        #-------------------------------------------
        repetition = cell.attributes["table:number-columns-repeated"]
        if(repetition)
          j = j+(repetition.to_i-1)
        end
        #-------------------------------------------
        # Falls Spaltenindex uebersprungen oder erreicht
        #-------------------------------------------
        if(j >= colInd)
          #-------------------------------------------
          # Zelltext und Datentyp zurueckgeben
          # ggf. Waehrungssymbol abschneiden
          #-------------------------------------------
          textElement = cell.elements["text:p"]
          if(! textElement)
            return nil,nil
          else
            text = textElement.text
            if(! text)
              text = ""
            end
            type = cell.attributes["office:value-type"]
            if(! type)
              type = "string"
            end
            text = normalize_text text,type
            return text,type
          end
        end
      }
      #----------------------------------------------
      # ausserhalb bisheriger Spalten
      #----------------------------------------------
      return nil,nil
    end
    ##########################################################################
    # Returns the content and type of the cell at the given indices
    # as strings. 
    # If the cell does not exist, nil is returned for text and type.
    # Type is one of the following office:value-types
    # * string, float, currency, time, date, percent, formula
    # The content of a formula is it's last calculated result or 0 in case of a
    # newly created cell. See annotations at 'readCellFromRow'.
    #   1.upto(10){ |i|
    #      text,type = readCell(i,i)
    #      write_cell(i,10-i,type,text)
    #   }
    #-------------------------------------------------------------------------
    def readCell(rowInd,colInd)
      #------------------------------------------------------------------
      # Fuer alle Zeilen
      #------------------------------------------------------------------
      i = 0
      j = 0
      #------------------------------------------------------------------
      # Zelle mit Indizes suchen
      #------------------------------------------------------------------
      currentTable = @tables[@current_table_name][NODE]
      currentTable.elements.each("table:table-row"){ |row|
        i = i+1  
        j = 0
        repetition = row.attributes["table:number-rows-repeated"]
        #-------------------------------------------
        # Zeilenwiederholungen addieren
        #-------------------------------------------
        if(repetition)
          i = i+(repetition.to_i-1)
        end
        #-------------------------------------------
        # Falls Zeilenindex uebersprungen oder erreicht
        #-------------------------------------------
        if(i >= rowInd)
          return readCellFromRow(row,colInd)
        end
      }
      #--------------------------------------------
      # ausserhalb bisheriger Zeilen
      #--------------------------------------------
      return nil,nil
    end
    ##########################################################################
    # internal: Composes everything necessary for writing all the contents of
    # the resulting *.ods zip-file upon call of 'save' or 'saveAs'.
    # Saves and zips all contents.
    #---------------------------------------
    def finalize zipfile
      initial_creator = @office_meta.elements["meta:initial-creator"]
      initial_creator = @office_meta.add_element REXML::Element.new "meta:initial-creator"
      die "Could not extract meta:initial-creator" unless initial_creator
      meta_creation_date = @office_meta.elements["meta:creation-date"]
      die "could not extract meta:creation-date" unless meta_creation_date
      time = "#{Time.now.year}-#{Time.now.month}-#{Time.now.day}T#{Time.now.hour}:#{Time.now.min}:#{Time.now.sec}"
      meta_creation_date.text = time
      meta_document_statistic = @office_meta.elements["meta:document-statistic"]
      die "Could not extract meta:document-statistic" unless meta_document_statistic
      meta_document_statistic.attributes["meta:table-count"] = @num_tables.to_s
      zipfile.file.open("meta.xml","w") { |outfile| outfile.puts @meta_text.to_s }
      zipfile.file.open("META-INF/manifest.xml","w") { |outfile| outfile.puts @manifest_text.to_s }
      zipfile.file.open("mimetype","w") { |outfile| outfile.print "application/vnd.oasis.opendocument.spreadsheet" }
      zipfile.file.open("settings.xml","w") { |outfile| outfile.puts @settings_text.to_s }
      zipfile.file.open("styles.xml","w") { |outfile| outfile.puts @styles_text.to_s }
      pad_tables
      zipfile.file.open("content.xml","w") { |outfile| outfile.puts @content_text.to_s }
    end
    ##########################################################################
    # internal: Called by constructor upon creation of Open Document-object.
    # Reads given zip-archive. Parses XML-files in archive. Initializes 
    # internal variables according to XML-trees. Calculates initial width of
    # all tables and creates default-styles and default-data-styles for all
    # data-types.
    #-------------------------------------------------------------------
    def init zipfile
      @meta_text = REXML::Document.new zipfile.file.read "meta.xml"
      @office_meta = @meta_text.elements["/office:document-meta/office:meta"]
      die"Could not extract office:document-meta" unless @office_meta
      @manifest_text = REXML::Document.new zipfile.file.read "META-INF/manifest.xml"
      @manifest_root = @manifest_text.elements["/manifest:manifest"]
      die "Could not extract manifest:manifest" unless @manifest_root
      @settings_text = REXML::Document.new zipfile.file.read "settings.xml"
      @office_settings = @settings_text.elements["/office:document-settings/office:settings"]
      die "Could not extract office:-settings" unless @office_settings
      @styles_text = REXML::Document.new zipfile.file.read "styles.xml"
      @office_styles = @styles_text.elements["/office:document-styles/office:styles"]
      die "Could not extract office:document-styles" unless @office_styles
      @content_text = REXML::Document.new zipfile.file.read "content.xml"
      @auto_styles = @content_text.elements["/office:document-content/office:automatic-styles"]
      die "Could not extract office:automatic-styles" unless @auto_styles
      init_house_keeping
      write_default_styles
    end
    ##########################################################################
    # internal: Converts the given string (of type 'float' or 'currency') to
    # the internal arithmetic represenation. 
    # This changes the thousands-separator, the decimal-separator and prunes
    # the currency-symbol
    #----------------------------------------------------------
    def normalize_text text, type
      new_text = String.new text
      if type == "currency" || type == "float"
        new_text.sub! /\./, ""
        new_text.sub! /,/, "."
        if type == "currency"
          new_text.sub! /\s*\S+$/, ""
        end
      end
      new_text
    end
    ##########################################################################
    # Writes the given text-string to given cell and sets style of
    # cell to corresponding type. Keep in mind: All values of tables are
    # passed and retrieved as strings
    #   sheet.write_text(sheet.get_cell(17,39),"currency","14,37")
    # The example can of course be simplified by
    #   sheet.write_cell 17,39,"currency","14,37"
    #-----------------------------------------------------------
    def write_text cell, type, text
      cell.attributes.each { |attribute,value| cell.attributes.delete attribute }
      if type == "string"
        cell.attributes["office:value-type"] = "string"
        cell.attributes["table:style-name"] = @string_style
      elsif type == "float"
        cell.attributes["office:value-type"] = "float"
        cell.attributes["office:value"] = text
        cell.attributes["table:style-name"] = @float_style
      elsif type.match /^formula/
        cell.attributes["table:formula"] = internalizeFormula text 
        case type
          when "formula","formula:float"
            cell.attributes["office:value-type"] = "float"
            cell.attributes["office:value"] = 0
            cell.attributes["table:style-name"] = @float_style
          when "formula:time"
            cell.attributes["office:value-type"] = "time"
            cell.attributes["office:time-value"] = "PT00H00M00S"
            cell.attributes["table:style-name"] = @time_style
          when "formula:date"
            cell.attributes["office:value-type"] = "date"
            cell.attributes["office:date-value"] = "0"
            cell.attributes["table:style-name"] = @date_style
          when "formula:currency"
            cell.attributes["office:value-type"] = "currency"
            cell.attributes["office:value"] = "0.0" # Recalculated when the file is opened
            cell.attributes["office:currency"] = @currency_symbol_internal
            cell.attributes["table:style-name"] = @currency_style
          else die("write_text: invalid type of formula #{type}")
        end
        text = "0"
      elsif type == "percent"
        cell.attributes["office:value-type"] = "percentage"
        cell.attributes["office:value"] = percent_to_percent_val text
        cell.attributes["table:style-name"] = @percent_style
        text = text+" %"
      elsif type == "currency"
        cell.attributes["office:value-type"] = "currency"
        cell.attributes["office:value"] = text
        text = "#{text} #{@currency_symbol}"
        cell.attributes["office:currency"] = @currency_symbol_internal
        cell.attributes["table:style-name"] = @currency_style
      elsif type == "date"
        cell.attributes["office:value-type"] = "date"
        cell.attributes["table:style-name"] = @date_style
        cell.attributes["office:date-value"] = date_to_date_val text
      elsif type == "time"
        cell.attributes["office:value-type"] = "time"
        cell.attributes["table:style-name"] = @time_style
        cell.attributes["office:time-value"] = time_to_time_val text
      else
        die "Wrong type #{type}"
      end
      if cell.elements["text:p"]
        cell.elements["text:p"].text = text
      else
        newElement = cell.add_element "text:p"
        newElement.text = text
      end
    end
    ##########################################################################
    # internal: Norms and maps a known set of attributes of the given style-Hash to
    # valid long forms of OASIS-style-attributes and replaces color-values with
    # their hex-representations.
    # Unknown hash-keys are copied as is.
    #-------------------------------------------------------------------------
    def norm_style_hash in_hash
      out_hash = Hash.new
      in_hash.each do |key,value|
        if key.match /^(fo:)?border/
          die "wrong format for border '#{value}'" unless value.match /^\S+\s+\S+\s+\S+$/
          # should match color out of "0.1cm solid red7"
          match = value.match /\S+\s\S+\s(\S+)\s*$/
          color = match[1]
          unless color.match /#[a-fA-F0-9]{6}/
            hex_color = Helpers.find_color_with_name color
            value.sub! color,hex_color
          end
        end
        case key
          when "name" then out_hash["style:name"] = value
          when "family" then out_hash["style:family"] = value
          when "parent-style-name" then out_hash["style:parent-style-name"] = value
          when "background-color" then out_hash["fo:background-color"] = value
          when "text-align-source" then out_hash["style:text-align-source"] = value
          when "text-align" then out_hash["fo:text-align"] = value
          when "margin-left" then out_hash["fo:margin-left"] = value
          when "color" then out_hash["fo:color"] = value
          when "border" then out_hash["fo:border"] = value
          when "border-bottom" then out_hash["fo:border-bottom"] = value
          when "border-top" then out_hash["fo:border-top"] = value
          when "border-left" then out_hash["fo:border-left"] = value
          when "border-right" then out_hash["fo:border-right"] = value
          when "font-style" then out_hash["fo:font-style"] = value
          when "font-weight" then out_hash["fo:font-weight"] = value
          when "data-style-name" then out_hash["style:data-style-name"] = value
          when "text-underline-style" then out_hash["style:text-underline-style"] = value
          when "text-underline-width" then out_hash["style:text-underline-width"] = value
          when "text-underline-color" then out_hash["style:text-underline-color"] = value
          else out_hash[key] = value
        end
      end
      out_hash
    end
    ##########################################################################
    # internal: Retrieves and returns the node of the style with the given name from content.xml or
    # styles.xml along with the indicator of the corresponding file.
    #-------------------------------------------------------------------------
    def get_style style_name
      style = @auto_styles.elements["*[@style:name = '#{styleName}']"]
      if style
        file = CONTENT
      else
        style = @office_styles.elements["*[@style:name = '#{styleName}']"]
        die "Could not find style \'#{styleName}\' in content.xml or styles.xml" unless style
        file = STYLES
      end
      return file, style
    end
    ##########################################################################
    # Merges style-attributes of given attribute-hash with current style
    # of given cell. Checks, whether the resulting style already exists in the
    # archive of created styles or creates and archives a new style. Applies the
    # found or created style to cell. Cell is a REXML::Element.
    #   sheet.setAttributes(cell,{ "border-right" => "0.05cm solid magenta4",
    #                                "border-bottom" => "0.03cm solid lightgreen",
    #                                "border-top" => "0.08cm solid salmon",
    #                                "font-style" => "italic",
    #                                "font-weight" => "bold"})
    #   sheet.setAttributes(cell,{ "border" => "0.01cm solid turquoise", # turquoise frame
    #                                "text-align" => "center",             # center alignment
    #                                "background-color" => "yellow2",      # background-color
    #                                "color" => "blue"})                   # font-color
    #   1.upto(7){ |row|
    #     cell = sheet.get_cell(row,5)
    #     sheet.setAttributes(cell,{ "border-right" => "0.07cm solid green6" }) 
    #   }
    #-------------------------------------------------------------------------
    def setAttributes(cell,attributes)
      die("setAttributes: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      die("setAttributes: hash #{attributes} is not a hash") unless (attributes.class.to_s == "Hash")
      #----------------------------------------------------------------------
      # Flag, ob neue Attribute und deren Auspraegungen bereits im aktuellen
      # style vorhanden sind
      #----------------------------------------------------------------------
      containsMatchingAttributes = TRUE
      #-----------------------------------------------------------------------
      # Attribut-Hash, welcher "convenience"-Werte enthalten kann (und wird ;-) 
      # zunaechst normieren
      #-----------------------------------------------------------------------
      attributes = norm_style_hash(attributes)
      die("setAttributes: attribute style:name not allowed in attribute-list as automatically generated") if (attributes.has_key?("style:name"))
      #------------------------------------------------------------------
      # Falls Zelle bereits style zugewiesen hat
      #------------------------------------------------------------------
      currentStyleName = cell.attributes["table:style-name"]
      if(currentStyleName)
        #---------------------------------------------------------------
        # style suchen (lassen)
        #---------------------------------------------------------------
        file,currentStyle = get_style(currentStyleName)
        #-----------------------------------------------------------------------
        # Pruefung, ob oben gefundener style die neuen Attribute und deren Werte
        # bereits enthaelt.
        # Falls auch nur ein Attribut nicht oder nicht mit dem richtigen Wert
        # vorhanden ist, muss ein neuer style erstellt werden.
        # Grundannahme: Ein Open-Document-Style-Attribut kann per se immer nur in einem bestimmten Typ
        # Knoten vorkommen und muss daher nicht naeher qualifiziert werden
        #-----------------------------------------------------------------------
        attributes.each{ |attribute,value|
          currentValue = currentStyle.attributes[attribute]
          #-------------------------------------------------
          # Attribut in Context-Node nicht gefunden ?
          #-------------------------------------------------
          if(! currentValue)  # nilClass
            # tell("setAttributes: #{currentStyleName}: #{attribute} not in Top-Node")
            #-----------------------------------------------------------
            # Attribut mit passendem Wert dann in Kind-Element vorhanden ?
            #-----------------------------------------------------------
            if(currentStyle.elements["*[@#{attribute} = '#{value}']"])
              # tell("setAttributes: #{currentStyleName}: #{attribute}/#{value} matching in Sub-Node")
            #-----------------------------------------------------------
            # andernfalls Komplettabbruch der Pruefschleife aller Attribute und Flag setzen
            # => neuer style muss erzeugt werden
            #-----------------------------------------------------------
            else
              # tell("setAttributes: #{currentStyleName}: #{attribute}/#{value} not matching in Sub-Node")
              containsMatchingAttributes = FALSE
              break
            end
          #--------------------------------------------------
          # Attribut in Context-Node gefunden
          #--------------------------------------------------
          else
            #--------------------------------------------------
            # Passt der Wert des gefundenen Attributes bereits ?
            #--------------------------------------------------
            if (currentValue == value)
              # tell("setAttributes: #{currentStyleName}: #{attribute}/#{value} matching in Top-Node")
            #-------------------------------------------------
            # bei unpassendem Wert Flag setzen
            #-------------------------------------------------
            else
              # tell("setAttributes: #{currentStyleName}: #{attribute}/#{value} not matching with #{currentValue} in Top-Node")
              containsMatchingAttributes = FALSE
            end
          end
        }
        #--------------------------------------------------------
        # Wurden alle Attribut-Wertepaare gefunden, d.h. kann 
        # bisheriger style weiterverwendet werden ?
        #--------------------------------------------------------
        if(containsMatchingAttributes)
        #-------------------------------------------------------
        # nein => passenden Style in Archiv suchen oder klonen und anpassen
        #-------------------------------------------------------
        else
          get_appropriate_style(cell,currentStyle,attributes)
        end
      #------------------------------------------------------------------------
      # Zelle hatte noch gar keinen style zugewiesen
      #------------------------------------------------------------------------
      else
        #----------------------------------------------------------------------
        # Da style fehlt, ggf. aus office:value-type bestmoeglichen style ermitteln
        #----------------------------------------------------------------------
        valueType = cell.attributes["office:value-type"]
        if(valueType)
          case valueType
            when "string" then currentStyleName = "string_style"
            when "percentage" then currentStyleName = "percent_styleage"
            when "currency" then currentStyleName = "currency_style"
            when "float" then currentStyleName = "float_style"
            when "date" then currentStyleName = "date_style"
            when "time" then currentStyleName = "time_style"
          else
            die("setAttributes: unknown office:value-type #{valueType} found in #{cell}")
          end
        else
          #-----------------------------------------
          # 'string_style' als Default
          #-----------------------------------------
          currentStyleName = "string_style" 
        end
        #-------------------------------------------------------
        # passenden Style in Archiv suchen oder klonen und anpassen
        #-------------------------------------------------------
        file,currentStyle = get_style(currentStyleName)
        get_appropriate_style(cell,currentStyle,attributes)
      end
    end
    ##########################################################################
    # internal: Function is called, when 'setAttributes' detected, that the current style
    # of a cell and a given attribute-list don't match. The function clones the current 
    # style of the cell, generates a virtual new style, merges it with the attribute-list,
    # calculates a hash-value of the resulting style, checks whether the latter is already 
    # in the pool of archived styles, retrieves an archived style or
    # writes the resulting new style, archives the latter and applies the effective style to cell.
    #-------------------------------------------------------------------------
    def get_appropriate_style cell, current_style, attributes
      if attributes.has_key? "style:name"
        die "attribute style:name not allowed in attribute-list as automatically generated"
      end
      new_style = clone_node current_style
      new_style_name = "auto_style#{@style_counter += 1}"
      attributes["style:name"] = new_style_name
      insert_style_attributes new_style, attributes
      hash_key = style_to_hash new_style
      if @style_archive.has_key? hashKey
        archive_style_name = @style_archive[hash_key]
        cell.attributes["table:style-name"] = archive_style_name
        @style_counter -= 1
        new_style = nil
      else
        @style_archive[hash_key] = new_style_name
        cell.attributes["table:style-name"] = new_style_name
        @auto_styles.elements << newStyle
      end
    end
    ##########################################################################
    # internal: verifies the validity of a hash of style-attributes.
    # The attributes have to be normed already.
    #-------------------------------------------------------------------------
    def check_style_attributes attributes
      attributes.each do |key,value|
        unless  key.match /:/
          die "found unnormed or invalid attribute #{key}"
        end
      end
      if attributes.has_key? "style:text-underline-style"
        attributes["style:text-underline-width"] = "auto" unless attributes.has_key? "style:text-underline-width"
        attributes["style:text-underline-color"] = "#000000" unless attributes.has_key? "style:text-underline-color"
      end
      if (attributes.has_key?("style:text-underline-width") || attributes.has_key?("style:text-underline-color")) &&  ! attributes.has_key?("style:text-underline-style")
        die "missing  style:text-underline-style ... please specify"
      end
      font_style = attributes["fo:font-style"]
      if font_style
        #if attributes.has_key?("fo:font-style-asian") || attributes.has_key?("fo:font-style-complex")
          # tell("automatically overwritten fo:font-style-asian/complex with value of fo:font-style")
        #end
        attributes["fo:font-style-asian"] = attributes["fo:font-style-complex"] = font_style
      end
      font_weight = attributes["fo:font-weight"]
      if font_weight
        #if attributes.has_key?("fo:font-weight-asian") || attributes.has_key?("fo:font-weight-complex")
          # tell "automatically overwritten fo:font-weight-asian/complex with value of fo:font-weight"
        #end
        attributes["fo:font-weight-asian"] = attributes["fo:font-weight-complex"] = font_weight
      end
      if attributes.has_key?("fo:border") \
         && (attributes.has_key?("fo:border-bottom") \
             || attributes.has_key?("fo:border-top") \
             || attributes.has_key?("fo:border-left") \
             || attributes.has_key?("fo:border-right"))
        # tell "automatically deleted fo:border as one or more sides were specified'"
        attributes.delete "fo:border"
      end
      left_margin = attributes["fo:margin-left"]
      text_align = attributes["fo:text-align"]
      if left_margin && text_align && text_align != "start" && left_margin != "0"
        # tell "automatically corrected: fo:text-align \'#{attributes['fo:text-align']}\' does not match fo:margin-left \'#{attributes['fo:margin-left']}\'"
        attributes["fo:margin-left"] = "0" 
      elsif left_margin && left_margin != "0" && !text_align
        # tell "automatically corrected: fo:margin-left \'#{attributes['fo:margin-left']}\' needs fo:text-align \'start\' to work"
        attributes["fo:text-align"] = "start" 
      end 
    end
    ##########################################################################
    # internal: Merges a hash of given style-attributes with those of
    # the given style-node. The attributes have to be normed already. Existing
    # attributes of the style-node are overwritten.
    #-------------------------------------------------------------------------
    def insert_style_attributes style, attributes
      die "Missing attribute style:name in node #{style}" unless style.attributes["style:name"]
      table_cell_properties = style.elements["style:table-cell-properties"]
      text_properties = style.elements["style:text-properties"]
      paragraph_properties = style.elements["style:paragraph-properties"]
      check_style_attributes attributes 
      attributes.each do |key,value|
        if key.match /^fo:border/ ||  key == "style:text-align-source" || key ==  "fo:background-color"
          table_cell_properties = style.add_element "style:table-cell-properties" unless table_cell_properties
          if key.match /^fo:border-/
            table_cell_properties.attributes.delete "fo:border"
          end
          table_cell_properties.attributes[key] = value
        else
          case key
            when "style:name", "style:family", "style:parent-style-name", "style:data-style-name"
              style.attributes[key] = value
            when "fo:color", "fo:font-style", "fo:font-style-asian", "fo:font-style-complex",
                 "fo:font-weight", "fo:font-weight-asian", "fo:font-weight-complex", "style:text-underline-style",
                 "style:text-underline-width", "style:text-underline-color"
              text_properties = style.add_element "style:text-properties" unless text_properties
              text_properties.attributes[key] = value
              if key == "fo:font-style"
                text_properties.attributes["fo:font-style-asian"] = text_properties.attributes["fo:font-style-complex"] = value
              elsif key == "fo:font-weight"
                text_properties.attributes["fo:font-weight-asian"] = text_properties.attributes["fo:font-weight-complex"] = value
              end
            when "fo:margin-left","fo:text-align"
              paragraph_properties = style.add_element "style:paragraph-properties" unless paragraph_properties
              paragraph_properties.attributes[key] = value
          else
              die "invalid or not implemented attribute #{key}"
          end
        end
      end
    end
    ##########################################################################
    # internal: Clones a given node recursively and returns the top-node as REXML::Element
    #-------------------------------------------------------------------------
    def clone_node node
      new_node = node.clone
      node.elements.each do |child|
        new_node.elements << clone_node(child)
      end
      new_node
    end
    ##########################################################################
    # Creates a new style out of the given attribute-hash with abbreviated and simplified syntax.
    #   sheet.write_style_abbr {"name" => "new_percent_style",        # <- style-name to be applied to a cell
    #                           "margin-left" => "0.3cm",
    #                           "text-align" => "start",
    #                           "color" => "blue",
    #                           "border" => "0.01cm solid black",
    #                           "font-style" => "italic",
    #                           "data-style-name" => "percent_format_style", # <- predefined RODS data-style
    #                           "font-weight" => "bold"}
    #-------------------------------------------------------------------------
    def write_style_abbr attributes
      write_style norm_style_hash attributes
    end
    ##########################################################################
    # internal: creates a style in content.xml out of the given attribute-hash, which has to be
    # supplied in fully qualified (normed) form. Missing attributes are replaced by default-values.
    #-------------------------------------------------------------------------
    def write_style attributes
      die "Missing attribute style:name" unless attributes.has_key? "style:name"
      table_cell_properties = Hash.new
      table_cell_properties[TAG] = "style:table-cell-properties"
      text_properties = Hash.new
      text_properties[TAG] = "style:text-properties"
      paragraph_properties = Hash.new
      paragraph_properties[TAG] = "style:paragraph-properties"
      style_attributes = {TAG => "style:style",
                       "style:family" => "table-cell",
                       "style:parent-style-name" => "Default"}
      check_style_attributes attributes 
      attributes.each do |key,value|
        case key
          when "style:name" then style_attributes["style:name"] = value
          when "style:family" then style_attributes["style:family"] = value
          when "style:parent-style-name" then style_attributes["style:parent-style-name"] = value
          when "style:data-style-name" then style_attributes["style:data-style-name"] = value
          when "fo:background-color" then table_cell_properties["fo:background-color"] = value
          when "style:text-align-source" then table_cell_properties["style:text-align-source"] = value
          when "fo:border-bottom" then table_cell_properties["fo:border-bottom"] = value
          when "fo:border-top" then table_cell_properties["fo:border-top"] = value
          when "fo:border-left" then table_cell_properties["fo:border-left"] = value
          when "fo:border-right" then table_cell_properties["fo:border-right"] = value
          when "fo:border" then table_cell_properties["fo:border"] = value
          when "fo:color" then text_properties["fo:color"] = value
          when "fo:font-style" then text_properties["fo:font-style"] = value
          when "fo:font-style-asian" then text_properties["fo:font-style-asian"] = value
          when "fo:font-style-complex" then text_properties["fo:font-style-complex"] = value
          when "fo:font-weight" then text_properties["fo:font-weight"] = value
          when "fo:font-weight-asian" then text_properties["fo:font-weight-asian"] = value
          when "fo:font-weight-complex" then text_properties["fo:font-weight-complex"] = value
          when "fo:margin-left" then paragraph_properties["fo:margin-left"] = value
          when "fo:text-align" then paragraph_properties["fo:text-align"] = value
        else
          die "invalid or not implemented attribute #{key}"
        end
      end
      style_attributes["child1"] = table_cell_properties if table_cell_properties.length > 1
      style_attributes["child2"] = text_properties if text_properties.length > 1
      style_attributes["child3"] = paragraph_properties if paragraph_properties.length > 1
      write_style_xml CONTENT, style_attributes
    end
    ##########################################################################
    # internal: write a style-XML-tree to content.xml or styles.xml. The given hash
    # has to be provided in qualified form. The new
    # style is archived in a hash-pool of styles. Prior to that the 'style:name' 
    # is replaced by a dummy-value to ensure comparability.
    # 
    # Caveat: RODS' default-styles cannot be overwritten
    # 
    # Example (internal setting of default date-style upon object creation)
    #    #------------------------------------------------------------------------
    #    # date
    #    #------------------------------------------------------------------------
    #    # date-Style part 1 (format)
    #    #--------------------------------------------------------
    #    write_style_xml(STYLES,{TAG => "number:date-style",
    #                   "style:name" => "date_format_style",
    #                   "number:automatic-order" => "true",
    #                   "number:format-source" => "language",
    #                   "child1" => {TAG => "number:day"},
    #                   "child2" => {TAG => "number:text",
    #                                TEXT => "."},
    #                   "child3" => {TAG => "number:month"},
    #                   "child4" => {TAG => "number:text",
    #                                TEXT => "."},
    #                   "child5" => {TAG => "number:year"}})
    #    #--------------------------------------------------------
    #    # date-Style part 2 (referencing format above)
    #    #--------------------------------------------------------
    #    write_style_xml(CONTENT,{TAG => "style:style",
    #                   "style:name" => "date_style",
    #                   "style:family" => "table-cell",
    #                   "style:parent-style-name" => "Default",
    #                   "style:data-style-name" => "date_format_style"})
    #------------------------------------------------------------------------
    def write_style_xml file, style_hash
      #-----------------------------------------------------------
      # Style with this name already exists? -> Delete,
      # If no default style of RODS, and from style-archive remove them.
      # style is only in the files of the two
      # Content.xml OR styles.xml wanted
      #-----------------------------------------------------------
      top_node = @auto_styles
      case file
        when STYLES then top_node = @office_styles
        when CONTENT then top_node = @auto_styles
        else die "write_style_xml: wrong file-parameter #{file}"
      end
      die "Missing attribute style:name" unless style_hash.has_key? "style:name"
      style_name = style_hash["style:name"]
      is_rods_style = @rods_styles.index style_name
      style_node = top_node.elements["*[@style:name = '#{style_name}']"]
      if style_node && !is_rods_style
        top_node.elements.delete style_node
        @style_archive.each do |key,value|
          if value == style_name
            @style_archive.delete key
            break
          end
        end
      end
      unless style_node && is_rods_style
        hashKey = style_to_hash write_xml top_node, style_hash
        @style_archive[hashKey] = style_name unless @style_archive.has_key? hashKey
      end
    end
    ##########################################################################
    # internal: converts XML-node of a style into a hash-value and returns
    # the string-representation of the latter.
    ##########################################################################
    def style_to_hash style_node
      style_node_string = style_node.to_s
      style_node_string.sub! /style:name\s* = \s*('|")\S+('|")/, "style:name = #{DUMMY}"
      style_node_string.gsub! /\s+/, ""
      sorted_string = style_node_string.split(//).sort.join
      sorted_string.hash.to_s
    end
    ##########################################################################
    # internal: write initial default styles into content.xml and styles.xml
    #------------------------------------------------------------------------
    def write_default_styles
      #------------------------------------------------------------------------
      # Formate fuer die Anlage von Tabellen
      #------------------------------------------------------------------------
      # Tabellenformat selbst
      #------------------------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                             "style:name" => "table_style",
                             "style:family" => "table",
                             "style:master-page-name" => "Default",
                             CHILD => {TAG => "style:table-properties",
                                       "style:writing-mode" => "lr-tb",
                                       "table:display" => "true"}})
      #------------------------------------------------------------------------
      # Zeilenformat 
      #------------------------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                             "style:name" => "row_style",
                             "style:family" => "table-row",
                             CHILD => {TAG => "style:table-row-properties",
                                       "style:use-optimal-row-height" => "true",
                                       "style:row-height" => "0.452cm",
                                       "fo:break-before" => "auto"}})
      #------------------------------------------------------------------------
      # Spaltenformat 
      #------------------------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                             "style:name" => "column_style",
                             "style:family" => "table-column",
                             CHILD => {TAG => "style:table-column-properties",
                                       "style:column-width" => "2.267cm",
                                       "style:row-height" => "0.452cm",
                                       "fo:break-before" => "auto"}})
      #------------------------------------------------------------------------
      # Float/Formula
      #------------------------------------------------------------------------
      # Float-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:number-style",
                     "style:name" => "float_format_style",
                     CHILD => {TAG => "number:number",
                               "number:decimal-places" => "2",
                               "number:min-integer-digits" => "1"}})
      #--------------------------------------------------------
      # Float-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "float_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "float_format_style"})
      #------------------------------------------------------------------------
      # Zeit
      #------------------------------------------------------------------------
      # Zeit-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:time-style",
                            "style:name" => "time_format_style_seconds",
                            "child1" => {TAG => "number:hours",
                                         "number:style" => "long"},
                            "child2" => {TAG => "number:text",
                                         TEXT => ":"},
                            "child3" => {TAG => "number:minutes",
                                         "number:style" => "long"},
                            "child4" => {TAG => "number:text",
                                         TEXT => ":"},
                            "child5" => {TAG => "number:seconds",
                                         "number:style" => "long"}})
      #--------------------------------------------------------
      # Zeit-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "time_seconds_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "time_format_style_seconds"})
      #------------------------------------------------------------------------
      # Zeit
      #------------------------------------------------------------------------
      # Zeit-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:time-style",
                            "style:name" => "time_format_style",
                            "child1" => {TAG => "number:hours",
                                         "number:style" => "long"},
                            "child2" => {TAG => "number:text",
                                         TEXT => ":"},
                            "child3" => {TAG => "number:minutes",
                                         "number:style" => "long"}})
      #--------------------------------------------------------
      # Zeit-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "time_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "time_format_style"})
      #------------------------------------------------------------------------
      # Prozent
      #------------------------------------------------------------------------
      # Prozent-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:percent-style",
                     "style:name" => "percent_format_style",
                     "child1" => {TAG => "number:number",
                                  "number:decimal-places" => "2",
                                  "number:min-integer-digits" => "1"},
                     "child2" => {TAG => "number:text",
                                  TEXT => "%"}})
      #--------------------------------------------------------
      # Prozent-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "percent_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "percent_format_style"})
      #------------------------------------------------------------------------
      # String
      #------------------------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "string_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default"})
      #------------------------------------------------------------------------
      # Datum
      #------------------------------------------------------------------------
      # Date-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:date-style",
                     "style:name" => "date_format_style",
                     "number:automatic-order" => "true",
                     "number:format-source" => "language",
                     "child1" => {TAG => "number:day"},
                     "child2" => {TAG => "number:text",
                                  TEXT => "."},
                     "child3" => {TAG => "number:month"},
                     "child4" => {TAG => "number:text",
                                  TEXT => "."},
                     "child5" => {TAG => "number:year"}})
      #--------------------------------------------------------
      # Date-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "date_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "date_format_style"})
      #------------------------------------------------------------------------
      # Datum als Wochentag
      #------------------------------------------------------------------------
      # Date-Style Teil 1 (Format)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:date-style",
                            "style:name" => "date_format_day_style",
                            CHILD => {TAG => "number:day-of-week"}})
      #--------------------------------------------------------
      # Date-Style Teil 2 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "date_day_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "date_format_day_style"})
      #------------------------------------------------------------------------
      # Waehrung
      #------------------------------------------------------------------------
      # Currency-Style Teil 1 (Mapping bei positiver Zahl)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:currency-style",
                     "style:name" => "currency_format_positive_style",
                     "child1" => {TAG => "number:number",
                                  "number:decimal-places" => "2",
                                  "number:min-integer-digits" => "1",
                                  "number:grouping" => "true"},
                     "child2" => {TAG => "number:text",
                                  TEXT => " "},
                     "child3" => {TAG => "number:currency-symbol",
                                  "number:language" => @language,
                                  "number:country" => @country,
                                  TEXT => @currency_symbol}})
      #--------------------------------------------------------
      # Currency-Style Teil 2 (Format mit Referenz zu Mapping)
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "number:currency-style",
                     "style:name" => "currency_format_style",
                     "child1" => {TAG => "style:text-properties",
                                  "fo:color" => "#ff0000"},
                     "child2" => {TAG => "number:text",
                                  TEXT => "-" },
                     "child3" => {TAG => "number:number",
                                  "number:decimal-places" => "2",
                                  "number:min-integer-digits" => "1",
                                  "number:grouping" => "true"},
                     "child4" => {TAG => "number:text",
                                  TEXT => " " },
                     "child5" => {TAG => "number:currency-symbol",
                                  "number:language" => @language,
                                  "number:country" => @country,
                                  TEXT => @currency_symbol },
                     "child6" => {TAG => "style:map",
                                  "style:condition" => "value()> = 0",
                                  "style:apply-style-name" => "currency_format_positive_style" }})
      #--------------------------------------------------------
      # Currency-Style Teil 3 (Referenz zu Format oben)
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "currency_style",
                     "style:family" => "table-cell",
                     "style:parent-style-name" => "Default",
                     "style:data-style-name" => "currency_format_style"})
      #--------------------------------------------------------
      # Annotation-Styles Teil 1
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "style:style",
                     "style:name" => "comment_paragraph_style",
                     "style:family" => "paragraph",
                     "child1" => {TAG => "style:paragraph-properties",
                                  "style:writing-mode" => "page",
                                  "style:text-autospace" => "none",
                                  "style:line-break" => "normal"},
                     "child2" => {TAG => "style:text-properties",
                                  "style:text-overline-mode" => "continuous",
                                  "fo:country" => @country,
                                  "style:country-asian" => "CN",
                                  "fo:font-size" => "10pt",
                                  "fo:font-weight" => "normal",
                                  "fo:text-shadow" => "none",
                                  "fo:hyphenate" => "false",
                                  "style:font-name-asian" => "DejaVu Sans",
                                  "style:font-style-asian" => "normal",
                                  "style:font-name-comlex" => "Lohit Hindi",
                                  "style:text-overline-style" => "none",
                                  "style:text-outline" => "false",
                                  "style:font-size-asian" => "10pt",
                                  "fo:language" => @language,
                                  "style:text-emphasize" => "none",
                                  "style:font-style-complex" => "normal",
                                  "style:text-line-through-style" => "none",
                                  "style:font-weight-complex" => "normal",
                                  "style:font-weight-asian" => "normal",
                                  "style:font-relief" => "home",
                                  "style:font-size-complex" => "10 pt",
                                  "style:language-asian" => "zh",
                                  "style-text-underline-mode" => "continuous",
                                  "style:country-complex" => "IN",
                                  "fo:font-style" => "normal",
                                  "style:text-line-through-mode" => "continuous",
                                  "style:text-overline-color" => "font-color",
                                  "style:text-underline-style" => "none",
                                  "style:language-complex" => "hi",
                                  "style:font-name" => "Arial"}})
      #--------------------------------------------------------
      # Annotation-Styles Teil 2
      #--------------------------------------------------------
      write_style_xml(STYLES,{TAG => "style:style",
                     "style:name" => "comment_text_style",
                     "style:family" => "text",
                     "child" => {TAG => "style:text-properties",
                                 "style:text-overline-mode" => "continuous",
                                 "fo:country" => @country,
                                 "style:country-asian" => "CN",
                                 "fo:font-size" => "10pt",
                                 "fo:font-weight" => "normal",
                                 "fo:text-shadow" => "none",
                                 "style:font-name-asian" => "DejaVu Sans",
                                 "style:font-style-asian" => "normal",
                                 "style:font-name-complex" => "Lohit Hindi",
                                 "style:text-overline-style" => "none",
                                 "style:text-outline" => "false",
                                 "style:font-size-asian" => "10pt",
                                 "fo:language" => @language,
                                 "style:text-emphasize" => "none",
                                 "style:font-style-complex" => "normal",
                                 "style:text-line-through-style" => "none",
                                 "style:font-weight-complex" => "normal",
                                 "style:font-weight-asian" => "normal",
                                 "style:font-relief" => "none",
                                 "style:font-size-complex" => "10pt",
                                 "style:language-asian" => "zh",
                                 "style:text-underline-mode" => "continuous",
                                 "style:country-complex" => "IN",
                                 "fo:font-style" => "normal",
                                 "style:text-line-through-mode" => "continuous",
                                 "style:text-overline-color" => "font-color",
                                 "style:text-underline-style" => "none",
                                 "style:language-complex" => "hi",
                                 "style:font-name" => "Arial"}})
      #--------------------------------------------------------
      # Annotation-Styles Teil 3
      #--------------------------------------------------------
      write_style_xml(CONTENT,{TAG => "style:style",
                     "style:name" => "comment_graphics_style",
                     "style:family" => "graphic",
                     CHILD => {TAG => "style:graphic-properties",
                               "fo:padding-right" => "0.1cm",
                               "draw:marker-start-width" => "0.2cm",
                               "draw:auto-grow-width" => "false",
                               "draw:marker-start-center" => "false",
                               "draw:shadow" => "hidden",
                               "draw:shadow-offset-x" => "0.1cm",
                               "draw:shadow-offset-y" => "0.1cm",
                               "draw:marker-start" => "Linienende_20_1",
                               "fo:padding-top" => "0.1cm",
                               "draw:fill" => "solid",
                               "draw:caption-escape-direction" => "auto",
                               "fo:padding-left" => "0.1cm",
                               "draw:fill-color" => "#ffffcc",
                               "draw:auto-grow-height" => "true",
                               "fo:padding-bottom" => "0.1cm"}})
    end
    ##########################################################################
    # Helper-Tool: Prints all styles of styles.xml in indented ASCII-notation
    #   sheet.printOfficeStyles()
    # * Lines starting with 'E' are Element-Tags
    # * Lines starting with 'A' are Attributes
    # * Lines starting with 'T' are Element-Text
    # Sample output:
    #   E: style:style
    #     A: style:name => "comment_graphics_style"
    #     A: style:family => "graphic"
    #     E: style:graphic-properties
    #       A: fo:padding-right => "0.1cm"
    #       A: draw:marker-start-width => "0.2cm"
    #       A: draw:auto-grow-width => "false"
    #       A: draw:marker-start-center => "false"
    #       A: draw:shadow => "hidden"
    #       A: draw:shadow-offset-x => "0.1cm"
    #       A: draw:shadow-offset-y => "0.1cm"
    #       A: draw:marker-start => "Linienende_20_1"
    #       A: fo:padding-top => "0.1cm"
    #       A: draw:fill => "solid"
    #       A: draw:caption-escape-direction => "auto"
    #       A: fo:padding-left => "0.1cm"
    #       A: draw:fill-color => "#ffffcc"
    #       A: draw:auto-grow-height => "true"
    #       A: fo:padding-bottom => "0.1cm"
    #-------------------------------------------------------------------------
    def printOfficeStyles()
      printStyles(@office_styles,"  ")
    end
    ##########################################################################
    # Helper-Tool: Prints all styles of content.xml in indented ASCII-notation
    #   sheet.printAutoStyles()
    # * Lines starting with 'E' are Element-Tags
    # * Lines starting with 'A' are Attributes
    # * Lines starting with 'T' are Element-Text
    # Sample output:
    #   E: number:date-style
    #     A: style:name => "date_format_style"
    #     A: number:automatic-order => "true"
    #     A: number:format-source => "language"
    #     E: number:day
    #     E: number:text
    #       T: "."
    #     E: number:month
    #     E: number:text
    #       T: "."
    #     E: number:year
    #-------------------------------------------------------------------------
    def printAutoStyles()
      printStyles(@auto_styles,"  ")
    end
    ##########################################################################
    # internal: Helper-Tool: Prints out all styles of given node in an indented ASCII-notation
    #------------------------------------------------------------------------
    def printStyles(startNode,indent)
      startNode.elements.each("*"){ |element|
        #------------------------------------------
        # Tag extrahieren (Standard-Tag-Zeichen nach '<')
        #------------------------------------------
        # puts("Element: #{element}")
        element.to_s.match(/<\s*([A-Za-z:-]+)/)
        puts("#{indent}E: #{$1}")
        #------------------------------------------
        # Attribute ausgeben
        #------------------------------------------
        element.attributes.each{ |attribute, value|
          puts("  #{indent}A: #{attribute} => \"#{value}\"")
        }
        #------------------------------------------
        # Text
        #------------------------------------------
        if(element.has_text?())
          puts("  #{indent}T: \"#{element.text}\"")
        end
        #------------------------------------------
        # Rekursion
        #------------------------------------------
        if(element.has_elements?())
          printStyles(element,indent+"  ")
        end
      }
    end
    ##########################################################################
    # internal: Recursively writes an XML-tree out of the given hash and returns
    # the written node. The returned node is irrelevant for the recursion but
    # valid for saving the node in a hash-pool for later style-comparisons.
    #------------------------------------------------------------------------
    def write_xml node, treeHash
      tag = ""
      text = ""
      attributes = Hash.new
      grand_children = Hash.new
      treeHash.each do |key,value|
        case key
          when TAG then tag = value
          when TEXT then text = value
          else  
            if(key.match(/child/))
              grand_children[key] = value
            else
              attributes[key] = value
            end
        end
      end
      die("Missing Tag for XML-Tree") unless (tag != "")
      child = node.add_element tag, attributes
      child.text = text unless text == ""
      grand_children.each{ |key,hash| write_xml child, hash }
      child
    end
    ##########################################################################
    # internal: Convert given formula to internal representation.
    # Example: " = E6-E5+0,27" => "of: = [.E6]+[.E5]+0.27"
    #------------------------------------------------------------------------
    def internalizeFormula(formulaIn)
      if !formulaIn.match /^ = /
        die "internalizeFormula: Formula #{formulaIn} does not begin with \' = \'"
      end
      formulaOut = String.new(formulaIn)
      #---------------------------------------------
      # Praefix setzen
      #---------------------------------------------
      formulaOut.sub!(/^ = /,"oooc: = ")
      #---------------------------------------------
      # Dezimaltrennzeichen ',' durch '.' in Zahlen ersetzen
      #---------------------------------------------
      formulaOut.gsub!(/(\d),(\d)/,"\\1.\\2") 
      #---------------------------------------------
      # Zellbezeichnerformat AABC3421 in [.AABC3421] wandeln
      #---------------------------------------------
      formulaOut.gsub!(/((\$?[A-Ta-z'.0-9][A-Ta-z' .0-9]*)\.)?(\$?[A-Za-z]+\$?\d+(:\$?[A-Za-z]+\$?\d+)?)/,"[\\2.\\3]")
    end
    ##########################################################################
    # convert column number to letters for usage in formulas
    # 1 => A
    # implementation is not ideal, consumes useless memory for large n :-)
    def column_number_to_address(n)
      raise ArgumentError, "n should be >= 1" if n < 1
      n -= 1 # column A should map to 1, arrays start counting from 0
      chars = []
      digits = 0
      while n >= chars.size
        digits += 1
        chars << (('A' * digits)..('Z' * digits)).to_a
        chars.flatten!
      end
      return chars[n]
    end
    ##########################################################################
    # Applies style of given name to given cell and overwrites all previous style-settings
    # of the latter including the former data-style
    #   sheet.write_style_abbr {"name" => "strange_style",
    #                           "text-align" => "right",
    #                           "data-style-name" => "currency_format_style" <- don't forget data-style
    #                           "border-left" => "0.01cm solid grey4"}
    #   sheet.set_style cell, "strange_style" # <- style-name has to exist
    #-------------------------------------------------------------------------
    def setStyle(cell,styleName)
      #-----------------------------------------------------------------------
      # Ist Style gueltig, d.h. in content.xml vorhanden ?
      #-----------------------------------------------------------------------
      die("setStyle: style \'#{styleName}\' does not exist") unless (@auto_styles.elements["*[@style:name = '#{styleName}']"])
      cell.attributes['table:style-name'] = styleName
    end
    ##########################################################################
    # Inserts an annotation field for the given cell. 
    # Caveat: When you make the annotation permanently visible in a subsequent
    # OpenOffice.org-session, the annotation will always be displayed in the upper
    # left corner of the sheet. The temporary display of the annotation is not 
    # affected however.
    #   sheet.writeComment(cell,"by Dr. Heinz Breinlinger (who else)")
    #------------------------------------------------------------------------
    def writeComment(cell,comment)
      die("writeComment: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      die("writeComment: comment #{comment} is not a string") unless (comment.class.to_s == "String")
      #--------------------------------------------
      # Ggf. alten Kommentar loeschen
      #--------------------------------------------
      cell.elements.delete("office:annotation")
      write_xml(cell,{TAG => "office:annotation",
                     "svg:x" => "4.119cm",
                     "draw:caption-point-x" => "-0.61cm",
                     "svg:y" => "0cm",
                     "draw:caption-point-y" => "0.011cm",
                     "draw:text-style-name" => "comment_paragraph_style",
                     "svg:height" => "0.596cm",
                     "draw:style-name" => "comment_graphics_style",
                     "svg:width" => "2.899cm",
                     "child1" => {TAG => "dc:date",
                                  TEXT => "2010-01-01T00:00:00"
                                 },                    
                     "child2" => {TAG => "text:p",
                                  "text:style-name" => "comment_paragraph_style",
                                  TEXT => comment
                                 }
                    })
    end
    ##########################################################################
    # internal: Helper-tool to extract a large amount of color-values and help
    # build a color-lookup-table.
    #-------------------------------------------------------------------------
    def getColorPalette()
      #------------------------------------------------
      # Automatic-Styles aus content.xml
      #------------------------------------------------
      styles = @content_text.elements["/office:document-content/office:automatic-styles"]
      currentTable = @tables[@current_table_name][NODE]
      currentTable.elements.each("//table:table-cell"){ |cell|
        textElement = cell.elements["text:p"]
        #-----------------------------
        # Zelle mit Text ?
        #-----------------------------
        if(textElement)
          text = textElement.text
          #-------------------------------
          # Ist Zelle Style zugewiesen ?
          #-------------------------------
          styleName = cell.attributes['table:style-name']
          if(styleName)
            #-------------------------------------
            # Style vorhanden ?
            #-------------------------------------
            style = styles.elements["style:style[@style:name = '#{styleName}']"]
            die("Could not find style #{styleName}") unless (style)
            #-------------------------------------
            # Properties-Element ebenfalls vorhanden ?
            #-------------------------------------
            properties = style.elements["style:table-cell-properties"]
            die("Could not find table-cell-properties for #{styleName}") unless (properties)
            #-------------------------------------
            # Nun noch Hintergrundfarbe extrahieren
            #-------------------------------------
            hexColor = properties.attributes["fo:background-color"]
            puts("\"#{text}\" => \"#{hexColor}\",")
          end
        end
      }
    end
    ##########################################################################
    # Saves the file associated with the current RODS-object.
    #   sheet.save()
    #-------------------------------------------------------------------------
    def save()
      die("save: internal error: @file is not set -> cannot save file") unless (@file && (! @file.empty?))
      die("save: this should not happen: file #{@file} is missing") unless (File.exists?(@file))
      Zip::ZipFile.open(@file){ |zipfile|
        finalize(zipfile) 
      } 
    end
    ##########################################################################
    # Saves the current content to a new destination/file.
    # Caveat: Thumbnails are not created (these are normally part of the *.ods-zip-file).
    #   sheet.saveAs("/home/heinz/Work/Example.ods")
    #-------------------------------------------------------------------------
    def saveAs(newFile)
      if(File.exists?(newFile))
        File.delete(newFile)
      end
      #--------------------------------------------------------
      # Datei anlegen
      #--------------------------------------------------------
      Zip::ZipFile.open(newFile,true){ |zipfile|
        ["Configurations2","META-INF","Thumbnails"].each{ |dir|
          zipfile.mkdir(dir)
          zipfile.file.chmod(0755,dir)
        }
        ["accelerator","floater","images","menubar","popupmenu","progressbar","statusbar","toolbar"].each{ |dir|
          subDir = "Configurations2/"+dir
          zipfile.mkdir(subDir)
          zipfile.file.chmod(0755,subDir)
        }
        finalize(zipfile) 
      }
    end
    ##########################################################################
    #
    #   sheet = Rods.new("/home/heinz/Work/Template.ods") 
    #   sheet = Rods.new("/home/heinz/Work/Template.ods",["de,"DE","€","EUR"]) 
    #   sheet = Rods.new("/home/heinz/Work/Another.ods",["us","US","$","DOLLAR"]) 
    #
    # "de","DE","€","EUR" are the default-settings for the language, country,
    # external and internal currency-symbol. All these values merely affect
    # currency-values and annotations (the latter though not visibly).
    #-------------------------------------------------------------------------
    def initialize options = {}
      default = { language: :us, country: :US, external_currency: :'$', internal_currency: :DOLLAR }
      default.merge! options
      @content_text
      @language = default[:language]
      @country = default[:country]
      @currency_symbol = default[:external_currency]
      @currency_symbol_internal = default[:internal_currency]
      @spread_sheet
      @styles_text
      @meta_text
      @office_meta
      @manifest_text
      @manifest_root
      @settings_text
      @office_settings
      @current_table_name
      @tables = Hash.new
      @num_tables
      @office_styles
      @auto_styles
      @float_style = "float_style"
      @date_style = "date_style"  
      @string_style = "string_style"
      @currency_style = "currency_style"
      @percent_style = "percent_style"
      @time_style = "time_style"
      @style_counter = 0
      @file
      @style_archive = Hash.new
      @rods_styles = [
        "table_style",
        "row_style",
        "column_style",
        "float_format_style",
        "float_style",
        "time_format_style", 
        "time_style",
        "time_seconds_style",
        "percent_format_style",
        "percent_style",
        "string_style",
        "date_format_style", 
        "date_style",
        "date_format_day_style",
        "date_day_style",
        "currency_format_positive_style", 
        "currency_format_style",
        "currency_style",
        "comment_paragraph_style",
        "comment_text_style", 
        "comment_graphics_style"
      ]
      open default[:file] if default.has_key? :file
    end
    ##########################################################################
    # Fast Routine to get the previous row, because XML-Parser does not have
    # to start from top-node of document to find row
    # Returns previous row as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #
    #------------------------------------------------------------------------
    def getPreviousExistentRow(row)
      #----------------------------------------------------------------------
      # table:table-row und table:table-column sind Siblings
      # Letztere duerfen jedoch NICHT zurueckgegeben werden
      #----------------------------------------------------------------------
      previousSibling = row.previous_sibling
      if(previousSibling && previousSibling.elements["self::table:table-row"])
        return previousSibling
      else
        return nil
      end
    end
    ##########################################################################
    # Fast Routine to get the next cell, because XML-Parser does not have
    # to start from top-node of row to find cell
    # Returns next cell as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def getNextExistentCell(cell)
      return cell.next_sibling
    end
    ##########################################################################
    # Fast Routine to get the previous cell, because XML-Parser does not have
    # to start from top-node of row to find cell
    # Returns previous cell as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def getPreviousExistentCell(cell)
      return cell.previous_sibling
    end
    ##########################################################################
    # Fast Routine to get the next row, because XML-Parser does not have
    # to start from top-node of document to find row
    # Returns next row as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def getNextExistentRow(row)
      return row.next_sibling
    end
    ##########################################################################
    # Finds all cells with content 'content' and returns them along with the
    # indices of row and column as an array of hashes.
    #   [{:cell => cell,
    #     :row  => rowIndex,
    #     :col  => colIndex},
    #    {:cell => cell,
    #     :row  => rowIndex,
    #     :col  => colIndex}]
    #  
    # Regular expressions for 'content' are allowed but must be enclosed in 
    # single (not double) quotes
    #
    # In case of no matches at all, an empty array is returned.
    # 
    # The following finds all occurences of a comma- or dot-separated number,
    # consisting of 1 digit before and 2 digits behind the decimal-separator.
    #
    # array = sheet.getCellsAndIndicesFor('\d{1}[.,]\d{2}')
    #
    # Keep in mind that the content of a call with a formula is not the formula, but the
    # current value of the computed result.
    #
    # Also consider that you have to search for the external (i.e. visible)
    # represenation of a cell's content, not it's internal computational value.
    # For instance, when looking for a currency value of 1525 (that is shown as
    # '1.525 EUR'), you'll have to code
    #
    #   result = sheet.getCellsAndIndicesFor('1[.,]525')
    #   result.each{ |cellHash|
    #     puts("Found #{cellHash[:cell] on #{cellHash[:row] - #{cellHash[:col]")
    #   }
    #-------------------------------------------------------------------------
    def getCellsAndIndicesFor(content)
      die("getCellsAndIndicesFor: 'content' is not of typ String") unless (content.class.to_s == "String")
      result = Array.new()
      i = 0
      #----------------------------------------------------------------
      # Alle Text-Nodes suchen
      #----------------------------------------------------------------
      @spread_sheet.elements.each("//table:table-cell/text:p"){ |textNode|
        text = textNode.text
        #---------------------------------------------------------
        # Zelle gefunden ?
        #
        # 'content' darf regulaerer Ausdruck sein, muss dann jedoch
        # in einfachen Hochkommata uebergeben werden
        #---------------------------------------------------------
        if(text && (text.match(/#{content}/)))
          result[i] = Hash.new() 
          #-----------------------------------------------------
          # Zelle und Zellenindex ermitteln
          #-----------------------------------------------------
          cell = textNode.elements["ancestor::table:table-cell"]
          unless (cell)
            die("getCellsAndIndicesFor: internal error: Could not extract parent-cell of textNode with #{content}") 
          end
          colIndex = get_index(cell)
          #-----------------------------------------------------
          # Zeile und Zeilenindex ermitteln
          #-----------------------------------------------------
          row = textNode.elements["ancestor::table:table-row"]
          unless (row)
            die("getCellsAndIndicesFor: internal error: Could not extract parent-row of textNode with #{content}") 
          end
          rowIndex = get_index(row)
          result[i][:cell] = cell
          result[i][:row] = rowIndex
          result[i][:col] = colIndex
          i += 1
        end
      }
      return result
    end

    def get_number_of_siblings node
      get_index_and_or_number node, NUMBER
    end

    def get_index node
      get_index_and_or_number node, INDEX
    end

    def get_index_and_number node
      get_index_and_or_number node, BOTH
    end
    ##########################################################################
    # internal: Calculates index (in the sense of spreadsheet, NOT XML) of
    # given element (row, cell or column as REXML::Element) within the 
    # corresponding parent-element (table or row) or the number of siblings
    # of the same kind or both - depending on the flag given.
    #
    # In case of flag 'BOTH' the method returns TWO values
    #
    # index = get_index_and_or_number row, INDEX # -> Line-number within table
    # num_columns = get_index_and_or_number column, NUMBER # number of columns
    # index, num_columns = get_index_and_or_number row, BOTH # Line-number and total number of lines
    #-------------------------------------------------------------------------
    def get_index_and_or_number node, flag
      die "invalid flag '#{flag}'" unless flag == NUMBER || flag == INDEX || flag == BOTH
      if node.elements["self::table:table-cell"]
        kind_of_self = "table:table-cell"
        kind_of_parent = "table:table-row"
        kind_of_repetition = "table:number-columns-repeated"
      elsif node.elements["self::table:table-column"]
        kind_of_self = "table:table-column"
        kind_of_parent = "table:table"
        kind_of_repetition = "table:number-columns-repeated"
      elsif node.elements["self::table:table-row"]
        kind_of_self = "table:table-row"
        kind_of_parent = "table:table"
        kind_of_repetition = "table:number-rows-repeated"
      else
        die "passed element '#{node}' is neither cell, nor row or column"
      end
      parent = node.elements["ancestor::#{kind_of_parent}"]
      unless parent
        die "Could not extract parent of #{node}" 
      end
      index = number = 0
      parent.elements.each kind_of_self do |child|
        number += 1
        if child == node
          if flag == INDEX
            return number
          elsif flag == BOTH
            index = number
          end
        elsif repetition = child.attributes[kind_of_repetition]
          number += repetition.to_i - 1
        end
      end
      if flag == INDEX
        die "Could not calculate number of element #{node}"
      elsif flag == NUMBER
        return number 
      else
        return index, number
      end
    end
    ##########################################################################
    # internal: Inserts a new header-column before the given header-column thereby 
    # shifting existing header-columns
    #-------------------------------------------------------------------------
    def insertColumnBeforeInHeader(column)
      die("insertColumnBeforeInHeader: column #{column} is not a REXML::Element") unless (column.class.to_s == "REXML::Element")
      newColumn = create_column(1)
      column.previous_sibling = newColumn
      #-----------------------------------------
      # bisherige Tabellenbreite überschritten ?
      #-----------------------------------------
      lengthOfHeader = get_number_of_siblings(column)
      if(lengthOfHeader > @tables[@current_table_name][WIDTH])
        @tables[@current_table_name][WIDTH] = lengthOfHeader
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      return newColumn
    end
    ##########################################################################
    # Delets the cell to the right of the given cell
    #
    #   cell = sheet.write_get_cell 4,7,"date","16.01.2011"
    #   sheet.deleteCellAfter(cell)
    #-------------------------------------------------------------------------
    def deleteCellAfter(cell)
      die("deleteCellAfter: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      #--------------------------------------------------------
      # Entweder Wiederholungsattribut der aktuellen Zelle
      # dekrementieren oder ggf. Wiederholungsattribut der
      # Folgezelle dekrementieren oder selbige loeschen
      #--------------------------------------------------------
      repetitions = cell.attributes["table:number-columns-repeated"]
      if(repetitions && repetitions.to_i > 1)
        cell.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
      else
        nextCell = cell.next_sibling
        die("deleteCellAfter: cell is already last cell in row") unless (nextCell)
        nextRepetitions = nextCell.attributes["table:number-columns-repeated"]
        if(nextRepetitions && nextRepetitions.to_i > 1)
          nextCell.attributes["table:number-columns-repeated"] = (nextRepetitions.to_i-1).to_s
        else
          row = cell.elements["ancestor::table:table-row"]
          unless (row)
            die("deleteCellAfter: internal error: Could not extract parent-row of cell #{cell}") 
          end
          row.elements.delete(nextCell)
        end
      end
    end
    ##########################################################################
    # Delets the row below the given row
    #
    #   row = sheet.get_row(11)
    #   sheet.deleteRowBelow(row)
    #-------------------------------------------------------------------------
    def deleteRowBelow(row)
      die("deleteRowBelow: row #{row} is not a REXML::Element") unless (row.class.to_s == "REXML::Element")
      #--------------------------------------------------------
      # Entweder Wiederholungsattribut der aktuellen Zeile
      # dekrementieren oder ggf. Wiederholungsattribut der
      # Folgezeile dekrementieren oder selbige loeschen
      #--------------------------------------------------------
      repetitions = row.attributes["table:number-rows-repeated"]
      if(repetitions && repetitions.to_i > 1)
        row.attributes["table:number-rows-repeated"] = (repetitions.to_i-1).to_s
      else
        nextRow = row.next_sibling
        die("deleteRowBelow: row #{row} is already last row in table") unless (nextRow)
        nextRepetitions = nextRow.attributes["table:number-rows-repeated"]
        if(nextRepetitions && nextRepetitions.to_i > 1)
          nextRow.attributes["table:number-rows-repeated"] = (nextRepetitions.to_i-1).to_s
        else
          table = row.elements["ancestor::table:table"]
          unless (table)
            die("deleteRowBelow: internal error: Could not extract parent-table of row #{row}") 
          end
          table.elements.delete(nextRow)
        end
      end
    end
    ##########################################################################
    # Delets the cell at the given index in the given row
    #
    #   row = sheet.get_row(8)
    #   sheet.deleteCell(row,9)
    #-------------------------------------------------------------------------
    def deleteCellFromRow(row,colInd)
      die("deleteCell: row #{row} is not a REXML::Element") unless (row.class.to_s == "REXML::Element")
      die("deleteCell: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("deleteCell: invalid index #{colInd}") unless (colInd > 0)
      cell = getCellFromRow(row,colInd+1)
      deleteCellBefore(cell)
    end
    ##########################################################################
    # Delets the given cell.
    #
    # 'cell' is a REXML::Element as returned by get_cell(cellInd).
    #
    # startCell = sheet.get_cell(34,1)
    # while(cell = sheet.getNextExistentCell(startCell))
    #   sheet.deleteCell2(cell)
    # end
    #-------------------------------------------------------------------------
    def deleteCell2(cell)
      die("deleteCell2: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      #-------------------------------------------------------------------
      # Entweder Wiederholungszahl dekrementieren oder Zelle loeschen
      #-------------------------------------------------------------------
      repetitions = cell.attributes["table:number-columuns-repeated"]
      if(repetitions && repetitions.to_i > 1)
        cell.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
        # tell("deleteCell2: decrementing empty cells")
      else
        row = cell.elements["ancestor::table:table-row"]
        unless (row)
          die("deleteCell2: internal error: Could not extract parent-row of cell #{cell}") 
        end
        row.elements.delete(cell)
        # tell("deleteCell2: deleting non-empty cell")
      end
    end
    ##########################################################################
    # Delets the given row.
    #
    # 'row' is a REXML::Element as returned by get_row(rowInd).
    #
    # startRow = sheet.get_row(12)
    # while(row = sheet.getNextExistentRow(startRow))
    #   sheet.deleteRow2(row)
    # end
    #-------------------------------------------------------------------------
    def deleteRow2(row)
      die("deleteRow2: row #{row} is not a REXML::Element") unless (row.class.to_s == "REXML::Element")
      #-------------------------------------------------------------------
      # Entweder Wiederholungszahl dekrementieren oder Zeile loeschen
      #-------------------------------------------------------------------
      repetitions = row.attributes["table:number-rows-repeated"]
      if(repetitions && repetitions.to_i > 1)
        row.attributes["table:number-rows-repeated"] = (repetitions.to_i-1).to_s
        # tell("deleteRow2: decrementing empty rows")
      else
        table = row.elements["ancestor::table:table"]
        unless (table)
          die("deleteRow2: internal error: Could not extract parent-table of row #{row}") 
        end
        table.elements.delete(row)
        # tell("deleteRow2: deleting non-empty row")
      end
    end
    ##########################################################################
    # Delets the row at the given index
    #
    #   sheet.deleteRow(7)
    #-------------------------------------------------------------------------
    def deleteRow(rowInd)
      die("deleteRow: index #{rowInd} is not a Fixnum/Integer") unless (rowInd.class.to_s == "Fixnum")
      die("deleteRow: invalid index #{rowInd}") unless (rowInd > 0)
      row = get_row(rowInd+1)
      deleteRowAbove(row)
    end
    ##########################################################################
    # Delets the cell at the given indices
    #
    #   sheet.deleteCell(7,9)
    #-------------------------------------------------------------------------
    def deleteCell(rowInd,colInd)
      die("deleteCell: index #{rowInd} is not a Fixnum/Integer") unless (rowInd.class.to_s == "Fixnum")
      die("deleteCell: invalid index #{rowInd}") unless (rowInd > 0)
      die("deleteCell: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("deleteCell: invalid index #{colInd}") unless (colInd > 0)
      row = get_row(rowInd)
      deleteCellFromRow(row,colInd)
    end
    ##########################################################################
    # Delets the row above the given row
    #
    #   row = sheet.get_row(5)
    #   sheet.deleteRowAbove(row)
    #-------------------------------------------------------------------------
    def deleteRowAbove(row)
      die("deleteRowAbove: row #{row} is not a REXML::Element") unless (row.class.to_s == "REXML::Element")
      #--------------------------------------------------------
      # Entweder Wiederholungsattribut der vorherigen Zeile
      # dekrementieren oder selbige loeschen
      #--------------------------------------------------------
      previousRow = row.previous_sibling
      die("deleteRowAbove: row is already first row in row") unless (previousRow)
      previousRepetitions = previousRow.attributes["table:number-rows-repeated"]
      if(previousRepetitions && previousRepetitions.to_i > 1)
        previousRow.attributes["table:number-rows-repeated"] = (previousRepetitions.to_i-1).to_s
      else
        table = row.elements["ancestor::table:table"]
        unless (table)
          die("deleteRowAbove: internal error: Could not extract parent-table of row #{row}") 
        end
        table.elements.delete(previousRow)
      end
    end
    ##########################################################################
    # Delets the cell to the left of the given cell
    #
    #   cell = sheet.write_get_cell 4,7,"formula:currency"," = A1+B2"
    #   sheet.deleteCellBefore(cell)
    #-------------------------------------------------------------------------
    def deleteCellBefore(cell)
      die("deleteCellBefore: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      #--------------------------------------------------------
      # Entweder Wiederholungsattribut der vorherigen Zelle
      # dekrementieren oder selbige loeschen
      #--------------------------------------------------------
      previousCell = cell.previous_sibling
      die("deleteCellBefore: cell is already first cell in row") unless (previousCell)
      previousRepetitions = previousCell.attributes["table:number-columns-repeated"]
      if(previousRepetitions && previousRepetitions.to_i > 1)
        previousCell.attributes["table:number-columns-repeated"] = (previousRepetitions.to_i-1).to_s
      else
        row = cell.elements["ancestor::table:table-row"]
        unless (row)
          die("deleteCellBefore: internal error: Could not extract parent-row of cell #{cell}") 
        end
        row.elements.delete(previousCell)
      end
    end
    ##########################################################################
    # Inserts a new cell before the given cell thereby shifting existing cells
    #   cell = sheet.get_cell(5,1)
    #   sheet.insertCellBefore(cell) # adds cell at beginning of row 5
    #-------------------------------------------------------------------------
    def insertCellBefore(cell)
      die("insertCellBefore: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      newCell = create_cell(1)
      cell.previous_sibling = newCell
      #-----------------------------------------
      # bisherige Tabellenbreite überschritten ?
      #-----------------------------------------
      lengthOfRow = get_number_of_siblings(cell)
      if(lengthOfRow > @tables[@current_table_name][WIDTH])
        @tables[@current_table_name][WIDTH] = lengthOfRow
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      return newCell
    end
    ##########################################################################
    # Inserts a new cell after the given cell thereby shifting existing cells
    #   cell = sheet.get_cell(4,7)
    #   sheet.insertCellAfter(cell)
    #-------------------------------------------------------------------------
    def insertCellAfter(cell)
      die("insertCellAfter: cell #{cell} is not a REXML::Element") unless (cell.class.to_s == "REXML::Element")
      newCell = create_cell(1)
      cell.next_sibling = newCell
      #-----------------------------------------------------------------------
      # etwaige Wiederholungen uebertragen
      #-----------------------------------------------------------------------
      repetitions = cell.attributes["table:number-columns-repeated"]
      if(repetitions)
        cell.attributes.delete("table:number-columns-repeated")
        newCell.next_sibling = create_cell(repetitions.to_i)
      end
      #-----------------------------------------
      # bisherige Tabellenbreite ueberschritten ?
      #-----------------------------------------
      lengthOfRow = get_number_of_siblings(cell)
      if(lengthOfRow > @tables[@current_table_name][WIDTH])
        @tables[@current_table_name][WIDTH] = lengthOfRow
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      return newCell
    end
    ##########################################################################
    # Inserts and returns a cell at the given index in the given row, 
    # thereby shifting existing cells.
    #
    #   row = sheet.get_row(5)
    #   cell = sheet.insertCellFromRow(row,17) 
    #-------------------------------------------------------------------------
    def insertCellFromRow(row,colInd)
      die("insertCell: row #{row} is not a REXML::Element") unless (row.class.to_s == "REXML::Element")
      die("insertCell: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("insertCell: invalid index #{colInd}") unless (colInd > 0)
      cell = getCellFromRow(row,colInd)
      return insertCellBefore(cell)
    end
    ##########################################################################
    # Inserts and returns a cell at the given index, thereby shifting existing cells.
    #
    #   cell = sheet.insertCell(4,17) 
    #-------------------------------------------------------------------------
    def insertCell(rowInd,colInd)
      die("insertCell: index #{rowInd} is not a Fixnum/Integer") unless (rowInd.class.to_s == "Fixnum")
      die("insertCell: invalid index #{rowInd}") unless (rowInd > 0)
      die("insertCell: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("insertCell: invalid index #{colInd}") unless (colInd > 0)
      cell = get_cell(rowInd,colInd)
      return insertCellBefore(cell)
    end
    ##########################################################################
    # Inserts and returns a row at the given index, thereby shifting existing rows
    #   row = sheet.insertRow(1) # inserts row above former row 1
    #-------------------------------------------------------------------------
    def insertRow(rowInd)
      die("insertRow: invalid rowInd #{rowInd}") unless (rowInd > 0)
      die("insertRow: rowInd #{rowInd} is not a Fixnum/Integer") unless (rowInd.class.to_s == "Fixnum")
      row = get_row(rowInd)
      return insertRowAbove(row)
    end
    ##########################################################################
    # Inserts a new row above the given row thereby shifting existing rows
    #   row = sheet.get_row(1)
    #   sheet.insertRowAbove(row)
    #-------------------------------------------------------------------------
    def insertRowAbove(row)
      newRow = create_row(1)
      row.previous_sibling = newRow
      return newRow
    end
    ##########################################################################
    # Inserts a new row below the given row thereby shifting existing rows
    #   row = sheet.get_row(8)
    #   sheet.insertRowBelow(row)
    #-------------------------------------------------------------------------
    def insertRowBelow(row)
      newRow = create_row(1)
      row.next_sibling = newRow
      #-----------------------------------------------------------------------
      # etwaige Wiederholungen uebertragen
      #-----------------------------------------------------------------------
      repetitions = row.attributes["table:number-rows-repeated"]
      if(repetitions)
        row.attributes.delete("table:number-rows-repeated")
        newRow.next_sibling = create_row(repetitions.to_i)
      end
      return newRow
    end
    ##########################################################################
    # Deletes the column at the given index
    #
    #   sheet.deleteColumn(8)
    #-------------------------------------------------------------------------
    def deleteColumn(colInd)
      die("deleteColumn: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("deleteColumn: invalid index #{colInd}") unless (colInd > 0)
      currentWidth = @tables[@current_table_name][WIDTH]
      die("deleteColumn: column-index #{colInd} is outside valid range/current table width") if (colInd > currentWidth)
      #-------------------------------------------------------------------
      # Entweder Wiederholungsattribut der fraglichen Spalte dekrementieren
      # oder selbige loeschen
      #-------------------------------------------------------------------
      currentTable = @tables[@current_table_name][NODE]
      column = get_child_by_index(currentTable,COLUMN,colInd)
      repetitions = column.attributes["table:number-columns-repeated"]
      if(repetitions && repetitions.to_i > 1)
        column.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
      else
        table = column.elements["ancestor::table:table"]
        unless (table)
          die("deleteColumn: internal error: Could not extract parent-table of column #{column}") 
        end
        table.elements.delete(column)
      end
      #-----------------------------------------------
      # Fuer alle existierenden Zeilen neue Zelle an
      # Spaltenposition einfuegen und dabei implizit
      # Tabellenbreite aktualisieren
      #-----------------------------------------------
      row = get_row(1)
      deleteCellFromRow(row,colInd)
      i = 1
      while(row = getNextExistentRow(row)) # fuer alle Zeilen ab der zweiten
        deleteCellFromRow(row,colInd)
        i += 1
      end 
    end
    ##########################################################################
    # Inserts a column at the given index, thereby shifting existing columns
    #   sheet.insertColumn(1) # inserts column before former column 1
    #-------------------------------------------------------------------------
    def insertColumn(colInd)
      die("insertColumn: index #{colInd} is not a Fixnum/Integer") unless (colInd.class.to_s == "Fixnum")
      die("insertColumn: invalid index #{colInd}") unless (colInd > 0)
      currentTable = @tables[@current_table_name][NODE]
      #-----------------------------------------------
      # Neuer Spalteneintrag im Header mit impliziter
      # Aktualisierung der Tabellenbreite
      #-----------------------------------------------
      column = get_child_by_index(currentTable,COLUMN,colInd)
      insertColumnBeforeInHeader(column)
      #-----------------------------------------------
      # Fuer alle existierenden Zeilen neue Zelle an
      # Spaltenposition einfuegen und dabei implizit
      # Tabellenbreite aktualisieren
      #-----------------------------------------------
      row = get_row(1)
      cell = get_child_by_index(row,CELL,colInd)
      insertCellBefore(cell)
      i = 1
      while(row = getNextExistentRow(row)) # fuer alle Zeilen ab der zweiten
        cell = get_child_by_index(row,CELL,colInd)
        insertCellBefore(cell)
        i += 1
      end 
    end
    ##########################################################################
    # internal: returns cell at index if existent, nil otherwise
    #   row = getRowIfExists(4)
    #   if(row)
    #     cell = getCellFromRowIfExists(row,7)
    #     unless(cell) .....
    #   end
    #-------------------------------------------------------------------------
    def getCellFromRowIfExists(row,colInd)
      return getElementIfExists(row,CELL,colInd)
    end
    ##########################################################################
    # internal: returns row at index if existent, nil otherwise
    #   if(sheet.getRowIfExists(4))
    #     ........
    #   end
    #-------------------------------------------------------------------------
    def getRowIfExists(rowInd)
      currentTable = @tables[@current_table_name][NODE]
      return getElementIfExists(currentTable,ROW,rowInd)
    end
    ##########################################################################
    # internal: examines, whether element of given type (row, cell, column) and index
    # exists or not.
    # Returns the element or nil if not existent.
    #-------------------------------------------------------------------------
    def getElementIfExists(parent,type,index)
      die("getElementIfExists: invalid type #{type}")
      die("getElementIfExists: parent is not a REXML::Element") unless (parent.class.to_s == "REXML::Element")
      die("getElementIfExists: index #{index} is not a Fixnum/Integer") unless (index.class.to_s == "Fixnum")
      die("getElementIfExists: invalid range for index #{index}") unless (index > 0)
      #--------------------------------------------------------------
      # Typabhaengige Vorbelegungen
      #--------------------------------------------------------------
      case type
        when CELL
          kindOfSelf = "table:table-cell"
          kindOfParent = "table:table-row"
          kindOfRepetition = "table:number-columns-repeated"
        when COLUMN
          kindOfSelf = "table:table-column"
          kindOfParent = "table:table"
          kindOfRepetition = "table:number-columns-repeated"
        when ROW
          kindOfSelf = "table:table-row"
          kindOfParent = "table:table"
          kindOfRepetition = "table:number-rows-repeated"
        else
          die("getElementIfExists: invalid type #{type}")
      end
      #--------------------------------------------------------------
      # Ist Kind-Element mit Index in Vater-Element vorhanden ?
      #--------------------------------------------------------------
      i = 0
      parent.elements.each(kindOfSelf){ |child|
        i += 1
        #----------------------------------------------------------
        # Index ueberschritten ? -> Ruecksprung mit nil
        # Index gefunden ? -> Rueckgabe des Elementes
        # sonst: etwaige Wiederholungen zaehlen
        #----------------------------------------------------------
        if (i > index)
          return nil
        elsif(i == index)
          return child
        elsif(repetition = child.attributes[kindOfRepetition])
          index += repetition.to_i-1 # '-1', da aktuelles Element ebenfalls als Wiederholung zaehlt
        end
      }
      #-------------------------------------------------------
      # Index liegt ausserhalb vorhandener Kind-Elemente
      #-------------------------------------------------------
      return nil
    end
    ##########################################################################
    # internal: Opens zip-file
    #-------------------------------------------------------------------------
    def open file
      die "file #{file} does not exist" unless File.exists? file
      Zip::ZipFile.open(file) { |zipfile| init zipfile }
      @file = file
    end

    public :set_date_format, :write_get_cell, :write_cell, :writeGetCellFromRow, :writeCellFromRow,
           :getCellFromRow, :get_cell, :get_row, :rename_table, :set_current_table,
           :insert_table, :delete_table, :readCellFromRow, :readCell, :setAttributes, :write_style_abbr,
           :setStyle, :printOfficeStyles, :printAutoStyles, :getNextExistentRow, :getPreviousExistentRow,
           :getNextExistentCell, :getPreviousExistentCell, :insertTableAfter, :insertTableBefore,
           :writeComment, :save, :saveAs, :initialize, :write_text, :getCellsAndIndicesFor,
           :insertRowBelow, :insertRowAbove, :insertCellBefore, :insertCellAfter, :insertColumn,
           :insertRow, :insertCell, :insertCellFromRow, :deleteCellBefore, :deleteCellAfter,
           :deleteCell, :deleteCellFromRow, :deleteRowAbove, :deleteRowBelow, :deleteRow,
           :deleteColumn, :deleteRow2, :deleteCell2

    private :die, :create_cell, :create_row, :get_child_by_index, :create_element, :set_repetition, :init_house_keeping,
            :get_table_width, :pad_tables, :time_to_time_val, :percent_to_percent_val, :date_to_date_val,
            :finalize, :init, :normalize_text, :norm_style_hash, :get_style, :get_index,
            :get_number_of_siblings, :get_index_and_or_number, :create_column,
            :get_appropriate_style, :check_style_attributes, :insert_style_attributes, :clone_node,
            :write_style, :write_style_xml, :style_to_hash, :write_default_styles, :write_xml,
            :internalizeFormula, :getColorPalette, :open, :printStyles, :insertTableBeforeAfter,
            :insertColumnBeforeInHeader, :getElementIfExists, :getRowIfExists, :getCellFromRowIfExists
  end
end
