describe Rods::Document do

  describe 'when intialized  with no arguments' do

    it 'should not throw an error' do
      Rods::Document.new
    end

    it 'should create an empty sheet to work with' do
      doc = Rods::Document.new 
      doc.current_table.wont_be_nil
    end

  end

end
