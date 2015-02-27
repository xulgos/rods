describe Rods::Document do

  def create_document_from_empty_sheet
    Rods::Document.new file: "#{File.dirname __FILE__}/empty.ods"
  end

  describe 'when intialized with no arguments' do

    it 'should not throw an error' do
      Rods::Document.new
    end

    it 'should create an empty sheet to work with' do
      skip
      doc = Rods::Document.new 
      doc.current_table.wont_be_nil
    end

  end

  describe 'when initialized with a filename' do

    it 'should set up the first sheet to work with' do
      create_document_from_empty_sheet.current_table.wont_be_nil
    end

  end

  describe 'rename_table' do

    it 'should be able to change the name of the current table' do
      doc = create_document_from_empty_sheet
      doc.rename_table 'Sheet1', 'New'
      doc.current_table.must_equal 'New'
    end

  end

  describe 'insert_table' do
    
    it 'should create a new table at the end of the list' do
      doc = create_document_from_empty_sheet
      doc.table_count.must_equal 1
      doc.insert_table 'New'
      doc.table_count.must_equal 2
    end

    describe 'after inserting it' do

      describe 'set_current_table' do

        it 'should change the current_table to a specified table' do
          doc = create_document_from_empty_sheet
          doc.insert_table 'new'
          doc.set_current_table 'new'
          doc.current_table.must_equal 'new'
        end

      end
    end
  end

  describe 'delete_table' do

    it 'should raise an exception if trying to delete current_table' do
      doc = create_document_from_empty_sheet
      doc.current_table.must_equal 'Sheet1'
      lambda { doc.delete_table 'Sheet1' }.must_raise Rods::RodsError
    end

    it 'should raise an exception if trying to delete a table that does not exists' do
      doc = create_document_from_empty_sheet
      lambda { doc.delete_table 'new' }.must_raise Rods::RodsError
    end

    it 'should remove a table' do
      doc = create_document_from_empty_sheet
      doc.insert_table 'new'
      doc.table_count.must_equal 2
      doc.delete_table 'new'
      doc.table_count.must_equal 1
    end

  end

  describe 'get_cell' do

    it 'should find a cell with its index' do
      doc = create_document_from_empty_sheet
      cell = doc.get_cell 1,1
      cell.wont_be_nil
    end

  end

  describe 'write_cell' do

    it 'should write a cell the value given' do
      doc = create_document_from_empty_sheet
      doc.write_cell 1,1, "string", "blah"
    end

    it 'should write to a cell diffrent types of values' do
      doc = create_document_from_empty_sheet
      doc.write_cell 1,1, "time", "13:37"
      doc.write_cell 1,2, "date", "12.01.2015"
      doc.write_cell 1,3, "float", "1.5"
      doc.write_cell 1,4, "currency", "24.20"
      doc.write_cell 1,5, "percent", "20"
      doc.write_cell 2,1, "formula", " = A3 + 1.2"
      doc.write_cell 2,1, "formula:time", " = A1 + 1"
      doc.write_cell 2,1, "formula:date", " = A2 + 1"
      doc.write_cell 2,1, "formula:float", " = A3 + 1.2"
      doc.write_cell 2,1, "formula:currency", " = A4 + 1.1"
    end

    it 'when given a type it does not recognize it raises an exception' do
      doc = create_document_from_empty_sheet
      lambda { doc.write_cell 1, 1, "blah", "something" }.must_raise Rods::RodsError
    end

  end

end
