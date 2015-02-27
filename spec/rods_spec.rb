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

end
