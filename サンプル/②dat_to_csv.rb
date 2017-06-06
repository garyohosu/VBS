# coding: windows-31j

require "Kconv"
require 'FileUtils'

SHIP_DATE = "20160414"

Dir::glob("**/ok.dat").each{|f|

	open(f,"r"){|f1|
		savefile=f.gsub("dat","csv")
		open(savefile,"w"){|f2|
			while not f1.eof
				#line = Kconv.tosjis(f1.readline)
				line = f1.readline
				if (line.split(/\s*,\s*/)[1] > SHIP_DATE) and (line.split(/\s*,\s*/)[1].length==8) then
					f2.print line
				end
			end
		}
	}
}
