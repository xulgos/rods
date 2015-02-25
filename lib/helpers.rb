module Rods
  module Helpers
    ##########################################################################
    # Helper-function: Print palette of implemented color-mappings
    #   mySheet.printColorMap()
    # generates ouput like ...
    #   "lightturquoise" => "#00ffff",
    #   "lightred" => "#ff0000",
    #   "lightmagenta" => "#ff00ff",
    #   "yellow" => "#ffff00",
    # you can use for 'setAttributes' and 'writeStyleAbbr'.
    #-------------------------------------------------------------------------
    def print_color_map
      puts("convenience color-mappings")
      puts("-----------------------------------------")
      Rods::Color.constants.each do |const|
        puts "\t#{const} -> #{Rods::Color.const_get const}"
      end
      puts("You can use the convenience keys in 'setAttribute' and 'writeStyleAbbr'")
      puts("for the attributes")
      puts("  border,border-bottom, border-top, border-left, border-right")
      puts("  background-color")
      puts("  color")
    end
  end
end
