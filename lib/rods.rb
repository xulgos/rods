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
    def write_get_cell row_index, column_index, type, text
      cell = get_cell row_index, column_index
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
    def writeGetCellFromRow(row,column_index,type,text)
      cell = get_cell_from_row row, column_index
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
    def writeCellFromRow(row,column_index,type,text)
      cell = get_cell_from_row row, column_index
      write_text(cell,type,text)
    end
    ##########################################################################
    # Returns the cell at the given index in the given row.
    # Cell and row are REXML::Elements.
    # The cell is created if it does not exist.
    #   row = sheet.get_row(15)
    #   cell = sheet.get_cell_from_row row, 17 # 17th cell of 15th row
    # Looks a bit strange compared to
    #   cell = sheet.get_cell(15,17)
    # but is considerably faster if you are operating on several cells of the
    # same row as after locating the first cell of the row the XML-Parser can start
    # from the node of the already found row instead of having to locate the
    # row over and over again.
    #-------------------------------------------------------------------------
    def get_cell_from_row row, column_index
      get_child_by_index row, CELL, column_index
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
    #   sheet.insert_table_before "table2", "table1"
    #-------------------------------------------------------------------------
    def insert_table_before relative_table_name, table_name
      insert_table_before_after relative_table_name, table_name, BEFORE
    end
    ##########################################################################
    # Inserts a table of the given name after the given spreadsheet and updates
    # the internal table-administration.
    #   sheet.insert_table_after "table1", "table2"
    #-------------------------------------------------------------------------
    def insert_table_after relative_table_name, table_name
      insert_table_before_after relative_table_name, table_name, AFTER
    end
    ##########################################################################
    # internal: Inserts a table of the given name before or after the given spreadsheet and updates
    # the internal table-administration. The default position is 'after'.
    #   sheet.insert_table_before_after("table1","table2",BEFORE)
    #-------------------------------------------------------------------------
    def insert_table_before_after relative_table_name, table_name, position = AFTER
      die "table '#{relative_table_name}' does not exist" unless @tables.has_key? relative_table_name
      die "table '#{table_name}' already exists" if @tables.has_key? table_name
      relative_table = @spread_sheet.elements["*[@table:name = '#{relative_table_name}']"]
      die "Could not locate existing table #{relative_table_name}" unless relative_table
      new_table = REXML::Element.new "table:table"
      new_table.add_attributes({"table:name" =>  table_name,
                               "table:print" => "false",
                               "table:style-name" => "table_style"})
      write_xml new_table, {TAG => "table:table-column",
                         "table:style" => "column_style",
                         "table:default-cell-style-name" => "Default"}
      write_xml newTable, {TAG => "table:table-row",
                         "table:style-name" => "row_style",
                         CHILD => {TAG => "table:table-cell"}}
      case position
        when BEFORE then @spread_sheet.insert_before relativeTable, newTable
        when AFTER then @spread_sheet.insert_after relativeTable, newTable
        else die "invalid parameter #{position}"
      end
      @tables[table_name] = Hash.new
      @tables[table_name][NODE] = new_table
      @tables[table_name][WIDTH] = get_table_width new_table
      @tables[table_name][WIDTHEXCEEDED] = false
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
    def readCellFromRow(row,column_index)
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
        if(j >= column_index)
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
    def readCell(row_index,column_index)
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
        if(i >= row_index)
          return readCellFromRow(row,column_index)
        end
      }
      #--------------------------------------------
      # ausserhalb bisheriger Zeilen
      #--------------------------------------------
      return nil,nil
    end
    ##########################################################################
    # internal: Composes everything necessary for writing all the contents of
    # the resulting *.ods zip-file upon call of 'save' or 'save_as'.
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
        cell.attributes["table:formula"] = internalize_formula text
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
    #   sheet.set_attributes cell, { "border-right" => "0.05cm solid magenta4",
    #                                "border-bottom" => "0.03cm solid lightgreen",
    #                                "border-top" => "0.08cm solid salmon",
    #                                "font-style" => "italic",
    #                                "font-weight" => "bold"})
    #   sheet.set_attributes cell, { "border" => "0.01cm solid turquoise", # turquoise frame
    #                                "text-align" => "center",             # center alignment
    #                                "background-color" => "yellow2",      # background-color
    #                                "color" => "blue"})                   # font-color
    #   1.upto(7) do |row|
    #     cell = sheet.get_cell row, 5
    #     sheet.set_attributes cell, { "border-right" => "0.07cm solid green6" }
    #   end
    #-------------------------------------------------------------------------
    def set_attributes cell, attributes
      contains_matching_attributes = true
      attributes = norm_style_hash attributes
      if attributes.has_key? "style:name"
        die "attribute style:name not allowed in attribute-list as automatically generated"
      end
      current_style_name = cell.attributes["table:style-name"]
      if current_style_name
        file, current_style = get_style current_style_name
        attributes.each do |attribute, value|
          current_value = current_style.attributes[attribute]
          if ! current_value
            unless current_style.elements["*[@#{attribute} = '#{value}']"]
              contains_matching_attributes = true
              break
            end
          else
            unless current_value == value
              contains_matching_attributes = false
            end
          end
        end
        unless contains_matching_attributes
          get_appropriate_style cell, current_style, attributes
        end
      else
        value_type = cell.attributes["office:value-type"]
        if value_type
          case value_type
            when "string" then current_style_name = "string_style"
            when "percentage" then current_style_name = "percent_style"
            when "currency" then current_style_name = "currency_style"
            when "float" then current_style_name = "float_style"
            when "date" then current_style_name = "date_style"
            when "time" then current_style_name = "time_style"
          else
            die "unknown office:value-type #{value_type} found in #{cell}"
          end
        else
          current_style_name = "string_style"
        end
        file, current_style = get_style current_style_name
        get_appropriate_style cell, current_style, attributes
      end
    end
    ##########################################################################
    # internal: Function is called, when 'set_attributes' detected, that the current style
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
    def internalize_formula formula_in
      unless formula_in.match /^ = /
        die "Formula #{formula_in} does not begin with \' = \'"
      end
      formula_out = String.new formula_in
      formula_out.sub! /^ = /,"oooc: = "
      formula_out.gsub! /((\$?[A-Ta-z'.0-9][A-Ta-z' .0-9]*)\.)?(\$?[A-Za-z]+\$?\d+(:\$?[A-Za-z]+\$?\d+)?)/,"[\\2.\\3]"
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
    def set_style cell, style_name
      die "set_style: style \'#{style_name}\' does not exist" unless @auto_styles.elements["*[@style:name = '#{style_name}']"]
      cell.attributes['table:style-name'] = style_name
    end
    ##########################################################################
    # Inserts an annotation field for the given cell.
    # Caveat: When you make the annotation permanently visible in a subsequent
    # OpenOffice.org-session, the annotation will always be displayed in the upper
    # left corner of the sheet. The temporary display of the annotation is not
    # affected however.
    #   sheet.write_comment cell, "this is a comment"
    #------------------------------------------------------------------------
    def write_comment cell, comment
      cell.elements.delete "office:annotation"
      write_xml cell, {TAG => "office:annotation",
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
                    }
    end
    ##########################################################################
    # Saves the file associated with the current RODS-object.
    #   sheet.save
    #-------------------------------------------------------------------------
    def save
      die "@file is not set -> cannot save file" if @file.nil? && @file.empty?
      die "file #{@file} is missing" unless File.exists? @file
      Zip::ZipFile.open(@file) { |zipfile| finalize zipfile }
    end
    ##########################################################################
    # Saves the current content to a new destination/file.
    # Caveat: Thumbnails are not created (these are normally part of the *.ods-zip-file).
    #   sheet.save_as "/home/heinz/Work/Example.ods"
    #-------------------------------------------------------------------------
    def save_as new_file
      if File.exists? new_file
        File.delete new_file
      end
      Zip::ZipFile.open(newFile, true) do |zipfile|
        ["Configurations2","META-INF","Thumbnails"].each do |dir|
          zipfile.mkdir dir
          zipfile.file.chmod 0755, dir
        end
        ["accelerator","floater","images","menubar","popupmenu","progressbar","statusbar","toolbar"].each do |dir|
          zipfile.mkdir "Configurations2/#{dir}"
          zipfile.file.chmod 0755, "Configurations2/#{dir}"
        end
        finalize zipfile
      end
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
    #------------------------------------------------------------------------
    def get_previous_existent_row row
      previous_sibling = row.previous_sibling
      if previous_sibling && previous_sibling.elements["self::table:table-row"]
        previous_sibling
      else
        nil
      end
    end
    ##########################################################################
    # Fast Routine to get the next cell, because XML-Parser does not have
    # to start from top-node of row to find cell
    # Returns next cell as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def get_next_existent_cell cell
      cell.next_sibling
    end
    ##########################################################################
    # Fast Routine to get the previous cell, because XML-Parser does not have
    # to start from top-node of row to find cell
    # Returns previous cell as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def get_previous_existent_cell cell
      cell.previous_sibling
    end
    ##########################################################################
    # Fast Routine to get the next row, because XML-Parser does not have
    # to start from top-node of document to find row
    # Returns next row as a REXML::Element or nil if no element exists.
    # Cf. explanation in README
    #------------------------------------------------------------------------
    def get_next_existent_row row
      row.next_sibling
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
    # array = sheet.get_cells_and_indices_for '\d{1}[.,]\d{2}'
    #
    # Keep in mind that the content of a call with a formula is not the formula, but the
    # current value of the computed result.
    #
    # Also consider that you have to search for the external (i.e. visible)
    # represenation of a cell's content, not it's internal computational value.
    # For instance, when looking for a currency value of 1525 (that is shown as
    # '1.525 EUR'), you'll have to code
    #
    #   result = sheet.get_cells_and_indices_for '1[.,]525'
    #   result.each do |cellHash|
    #     puts "Found #{cellHash[:cell] on #{cellHash[:row] - #{cellHash[:col]"
    #   end
    #-------------------------------------------------------------------------
    def get_cells_and_indices_for content
      result = Array.new
      i = 0
      @spread_sheet.elements.each("//table:table-cell/text:p") do |text_node|
        text = text_node.text
        if text && text.match(content)
          result[i] = Hash.new
          cell = text_node.elements["ancestor::table:table-cell"]
          unless cell
            die "Could not extract parent-cell of text_node with #{content}"
          end
          col_index = get_index cell
          row = text_node.elements["ancestor::table:table-row"]
          unless row
            die "Could not extract parent-row of text_node with #{content}"
          end
          row_index = get_index row
          result[i][:cell] = cell
          result[i][:row] = row_index
          result[i][:col] = col_index
          i += 1
        end
      end
      result
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
    def insert_column_before_in_header column
      newColumn = create_column 1
      column.previous_sibling = newColumn
      length_of_header = get_number_of_siblings column
      if length_of_header > @tables[@current_table_name][WIDTH]
        @tables[@current_table_name][WIDTH] = length_of_header
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      new_column
    end
    ##########################################################################
    # Delets the cell to the right of the given cell
    #
    #   cell = sheet.write_get_cell 4, 7, "date", "16.01.2011"
    #   sheet.delete_cell_after cell
    #-------------------------------------------------------------------------
    def delete_cell_after cell
      repetitions = cell.attributes["table:number-columns-repeated"]
      if repetitions && repetitions.to_i > 1
        cell.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
      else
        next_cell = cell.next_sibling
        die "cell is already last cell in row" unless next_cell
        next_repetitions = next_cell.attributes["table:number-columns-repeated"]
        if next_repetitions && next_repetitions.to_i > 1
          next_cell.attributes["table:number-columns-repeated"] = (next_repetitions.to_i-1).to_s
        else
          row = cell.elements["ancestor::table:table-row"]
          unless row
            die "Could not extract parent-row of cell #{cell}"
          end
          row.elements.delete next_cell
        end
      end
    end
    ##########################################################################
    # Delets the row below the given row
    #
    #   row = sheet.get_row 11
    #   sheet.delete_row_below row
    #-------------------------------------------------------------------------
    def delete_row_below row
      repetitions = row.attributes["table:number-rows-repeated"]
      if repetitions && repetitions.to_i > 1
        row.attributes["table:number-rows-repeated"] = (repetitions.to_i-1).to_s
      else
        next_row = row.next_sibling
        die "row #{row} is already last row in table" unless next_row
        next_repetitions = next_row.attributes["table:number-rows-repeated"]
        if next_repetitions && next_repetitions.to_i > 1
          next_row.attributes["table:number-rows-repeated"] = (next_repetitions.to_i-1).to_s
        else
          table = row.elements["ancestor::table:table"]
          unless table
            die "Could not extract parent-table of row #{row}"
          end
          table.elements.delete next_row
        end
      end
    end
    ##########################################################################
    # Delets the cell at the given index in the given row
    #
    #   row = sheet.get_row 8
    #   sheet.delete_cell row, 9
    #-------------------------------------------------------------------------
    def delete_cell_from_row row, column_index
      die "invalid index #{column_index}" unless  column_index > 0
      cell = get_cell_from_row row, column_index+1
      delete_cell_before cell
    end
    ##########################################################################
    # Delets the given cell.
    #
    # 'cell' is a REXML::Element as returned by get_cell cell_ind.
    #
    # start_cell = sheet.get_cell 34,1
    # while cell = sheet.get_next_existent_cell start_cell
    #   sheet.delete_cell_element cell
    # end
    #-------------------------------------------------------------------------
    def delete_cell_element cell
      repetitions = cell.attributes["table:number-columns-repeated"]
      if repetitions && repetitions.to_i > 1
        cell.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
      else
        row = cell.elements["ancestor::table:table-row"]
        unless row
          die "Could not extract parent-row of cell #{cell}"
        end
        row.elements.delete cell
      end
    end
    ##########################################################################
    # Deletes the given row.
    #
    # 'row' is a REXML::Element as returned by get_row row_index.
    #
    # start_row = sheet.get_row 12
    # while row = sheet.get_next_existent_row start_row
    #   sheet.delete_row_element row
    # end
    #-------------------------------------------------------------------------
    def delete_row_element row
      repetitions = row.attributes["table:number-rows-repeated"]
      if repetitions && repetitions.to_i > 1
        row.attributes["table:number-rows-repeated"] = repetitions.to_i-1.to_s
      else
        table = row.elements["ancestor::table:table"]
        unless table
          die "Could not extract parent-table of row #{row}"
        end
        table.elements.delete row
      end
    end
    ##########################################################################
    # Delets the row at the given index
    #
    #   sheet.delete_row 7
    #-------------------------------------------------------------------------
    def delete_row row_index
      die "invalid index #{row_index}" unless  row_index > 0
      row = get_row row_index + 1
      delete_row_above row
    end
    ##########################################################################
    # Delets the cell at the given indices
    #
    #   sheet.delete_cell 7, 9
    #-------------------------------------------------------------------------
    def delete_cell row_index, column_index
      die "invalid index #{row_index}" unless row_index > 0
      die "invalid index #{column_index}" unless column_index > 0
      row = get_row row_index
      delete_cell_from_row row, column_index
    end
    ##########################################################################
    # Delets the row above the given row
    #
    #   row = sheet.get_row 5
    #   sheet.delete_row_above row
    #-------------------------------------------------------------------------
    def delete_row_above row
      previous_row = row.previous_sibling
      die "row is already first row in row" unless previous_row
      previous_repetitions = previous_row.attributes["table:number-rows-repeated"]
      if previousRepetitions && previousRepetitions.to_i > 1
        previous_row.attributes["table:number-rows-repeated"] = (previous_repetitions.to_i-1).to_s
      else
        table = row.elements["ancestor::table:table"]
        unless table
          die "Could not extract parent-table of row #{row}"
        end
        table.elements.delete previous_row
      end
    end
    ##########################################################################
    # Delets the cell to the left of the given cell
    #
    #   cell = sheet.write_get_cell 4, 7, "formula:currency", " = A1+B2"
    #   sheet.delete_cell_before cell
    #-------------------------------------------------------------------------
    def delete_cell_before cell
      previous_cell = cell.previous_sibling
      die "cell is already first cell in row" unless previous_cell
      previous_repetitions = previous_cell.attributes["table:number-columns-repeated"]
      if previous_repetitions && previous_repetitions.to_i > 1
        previous_cell.attributes["table:number-columns-repeated"] = (previous_repetitions.to_i-1).to_s
      else
        row = cell.elements["ancestor::table:table-row"]
        unless row
          die "Could not extract parent-row of cell #{cell}"
        end
        row.elements.delete previous_cell
      end
    end
    ##########################################################################
    # Inserts a new cell before the given cell thereby shifting existing cells
    #   cell = sheet.get_cell 5, 1
    #   sheet.insert_cell_before cell # adds cell at beginning of row 5
    #-------------------------------------------------------------------------
    def insert_cell_before cell
      new_cell = create_cell 1
      cell.previous_sibling = new_cell
      length_of_row = get_number_of_siblings cell
      if length_of_row > @tables[@current_table_name][WIDTH]
        @tables[@current_table_name][WIDTH] = length_of_row
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      new_cell
    end
    ##########################################################################
    # Inserts a new cell after the given cell thereby shifting existing cells
    #   cell = sheet.get_cell 4, 7
    #   sheet.insert_cell_after cell
    #-------------------------------------------------------------------------
    def insert_cell_after cell
      new_cell = create_cell 1
      cell.next_sibling = new_cell
      repetitions = cell.attributes["table:number-columns-repeated"]
      if repetitions
        cell.attributes.delete "table:number-columns-repeated"
        new_cell.next_sibling = create_cell repetitions.to_i
      end
      length_of_row = get_number_of_siblings cell
      if length_of_row > @tables[@current_table_name][WIDTH]
        @tables[@current_table_name][WIDTH] = length_of_row
        @tables[@current_table_name][WIDTHEXCEEDED] = true
      end
      new_cell
    end
    ##########################################################################
    # Inserts and returns a cell at the given index in the given row,
    # thereby shifting existing cells.
    #
    #   row = sheet.get_row 5
    #   cell = sheet.insert_cell_from_row row, 17
    #-------------------------------------------------------------------------
    def insert_cell_from_row row, column_index
      die "insert_cell: invalid index #{column_index}" unless  column_index > 0
      cell = get_cell_from_row row, column_index
      insert_cell_before cell
    end
    ##########################################################################
    # Inserts and returns a cell at the given index, thereby shifting existing cells.
    #
    #   cell = sheet.insert_cell 4, 17
    #-------------------------------------------------------------------------
    def insert_cell row_index, column_index
      die "invalid index #{row_index}" unless row_index > 0
      die "invalid index #{column_index}" unless column_index > 0
      cell = get_cell row_index, column_index
      insert_cell_before cell
    end
    ##########################################################################
    # Inserts and returns a row at the given index, thereby shifting existing rows
    #   row = sheet.insert_row 1 # inserts row above former row 1
    #-------------------------------------------------------------------------
    def insert_row row_index
      die "invalid row_index #{row_index}" unless row_index > 0
      row = get_row row_index
      insert_row_above row
    end
    ##########################################################################
    # Inserts a new row above the given row thereby shifting existing rows
    #   row = sheet.get_row 1
    #   sheet.insert_row_above row
    #-------------------------------------------------------------------------
    def insert_row_above row
      new_row = create_row 1
      row.previous_sibling = new_row
      new_row
    end
    ##########################################################################
    # Inserts a new row below the given row thereby shifting existing rows
    #   row = sheet.get_row 8
    #   sheet.insert_row_below row
    #-------------------------------------------------------------------------
    def insert_row_below row
      new_row = create_row 1
      row.next_sibling = new_row
      repetitions = row.attributes["table:number-rows-repeated"]
      if repetitions
        row.attributes.delete "table:number-rows-repeated"
        new_row.next_sibling = create_row repetitions.to_i
      end
      new_row
    end
    ##########################################################################
    # Deletes the column at the given index
    #
    #   sheet.delete_column 8
    #-------------------------------------------------------------------------
    def delete_column column_index
      die "invalid index #{column_index}" unless  column_index > 0
      current_width = @tables[@current_table_name][WIDTH]
      die "column-index #{column_index} is outside valid range/current table width" if column_index > current_width
      current_table = @tables[@current_table_name][NODE]
      column = get_child_by_index current_table, COLUMN, column_index
      repetitions = column.attributes["table:number-columns-repeated"]
      if repetitions && repetitions.to_i > 1
        column.attributes["table:number-columns-repeated"] = (repetitions.to_i-1).to_s
      else
        table = column.elements["ancestor::table:table"]
        unless table
          die "Could not extract parent-table of column #{column}"
        end
        table.elements.delete column
      end
      row = get_row 1
      delete_cell_from_row row, column_index
      i = 1
      while row = get_next_existent_row(row)
        delete_cell_from_row row, column_index
        i += 1
      end
    end
    ##########################################################################
    # Inserts a column at the given index, thereby shifting existing columns
    #   sheet.insert_column 1 # inserts column before former column 1
    #-------------------------------------------------------------------------
    def insert_column column_index
      die "invalid index #{column_index}" unless column_index > 0
      current_table = @tables[@current_table_name][NODE]
      column = get_child_by_index current_table, COLUMN, column_index
      insert_column_before_in_header column
      row = get_row 1
      cell = get_child_by_index row, CELL, column_index
      insert_cell_before cell
      i = 1
      while row = get_next_existent_row(row)
        cell = get_child_by_index row, CELL, column_index
        insert_cell_before cell
        i += 1
      end
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
           :get_cell_from_row, :get_cell, :get_row, :rename_table, :set_current_table,
           :insert_table, :delete_table, :readCellFromRow, :readCell, :set_attributes, :write_style_abbr,
           :set_style, :get_next_existent_row, :get_previous_existent_row,
           :get_next_existent_cell, :get_previous_existent_cell, :insert_table_after, :insert_table_before,
           :write_comment, :save, :save_as, :initialize, :write_text, :get_cells_and_indices_for,
           :insert_row_below, :insert_row_above, :insert_cell_before, :insert_cell_after, :insert_column,
           :insert_row, :insert_cell, :insert_cell_from_row, :delete_cell_before, :delete_cell_after,
           :delete_cell, :delete_cell_from_row, :delete_row_above, :delete_row_below, :delete_row,
           :delete_column, :delete_row_element, :delete_cell_element

    private :die, :create_cell, :create_row, :get_child_by_index, :create_element, :set_repetition, :init_house_keeping,
            :get_table_width, :pad_tables, :time_to_time_val, :percent_to_percent_val, :date_to_date_val,
            :finalize, :init, :normalize_text, :norm_style_hash, :get_style, :get_index,
            :get_number_of_siblings, :get_index_and_or_number, :create_column,
            :get_appropriate_style, :check_style_attributes, :insert_style_attributes, :clone_node,
            :write_style, :write_style_xml, :style_to_hash, :write_default_styles, :write_xml,
            :internalize_formula, :open, :insert_table_before_after, :insert_column_before_in_header
  end
end
