describe Rods::Document do

  describe 'when intialized with no arguments' do

    it 'should not throw an error' do
      Rods::Document.new
    end

    it 'should create an empty sheet to work with' do
      doc = Rods::Document.new 
      doc.current_table.wont_be_nil
    end

  end

  describe 'when initialized with a filename' do

    it 'should set up the first sheet to work with' do
      doc = Rods::Document.new file: "#{File.dirname __FILE__}/empty.ods"
      doc.current_table.wont_be_nil
    end

  end

  describe 'rename_table' do

    it 'should be able to change the name of the current table' do
      doc = Rods::Document.new file: "#{File.dirname __FILE__}/empty.ods"
      doc.rename_table 'Sheet1', 'New'
      doc.current_table.must_equal 'New'
    end

  end

end
