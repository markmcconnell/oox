#!/usr/bin/ruby
$:.push('../lib')
require 'ox'

require 'fileutils'
FileUtils.rm_f('test.xlsx')

# create a new spreadsheet and write to file ad the end of the block
Ox::Spreadsheet.write('test.xlsx') {|ss|
	# create a new worksheet/tab inside the spreadsheet
	ss.worksheet("People") {|ws|
		# create a new row
		ws.row {
			# indicate that we're starting a table
			ws.table
			
			# these will be the "header" columns
			ws.cell('Name')
			ws.cell('Birth')
			ws.cell('Death')
			ws.cell('Occupation')
			
			# we could call this manually,  but it
			# is called automatically at the end of
			# the row:
			# ws.header!
		}
		
		ws.row {
			# _add_ to the default style for the worksheet
			ws.style(:bold => true)
			
			# new default style applied
			ws.cell('Mark McConnell')
			
			# _add_ to the default style for a single cell
			# this will be both bold and green
			ws.cell(1979, { :bg => '00CC00' })
			
			# but it's temporary,  this cell is only bold
			ws.cell('Isn\'t!')
			
			# the default style is a set,  so this only
			# changes a single attribute
			ws.style(:bold => false)
			
			# final cell,  which is back to default
			ws.cell('Programmer')
		}
		
		# if you don't need to deviate from the default style,  or
		# the whole row is a single style,  you can just add it 
		# in one call:
		ws.row('Wolfgang Mozart', 1756, 1791, 'Composer')
		
		ws.row {
			ws.cell('Sigmund Freud')
			
			# we also provide (cells) which allows you to insert
			# several cell values with the same style
			ws.cells(1856, 1939)
			
			ws.cell('Psychoanalyst')
		}
		
		# we can do this manually,  but it will happen explicity
		# due to the end of the worksheet block
		# ws.table!
	}
}

##############################################################################
__END__
