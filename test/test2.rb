#!/usr/bin/ruby
$:.push('../lib')
require 'ox'

ss = Ox::Spreadsheet.new()
ws = ss.worksheet("Beginning")

ws.row.table
ws.cell('Name')
ws.cell('Birth')
ws.cell('Death')
ws.cell('Occupation')
		
ws.row.cell('Mark McConnell').cell(1979).cell('--', { :bg => '00CC00' }).cell('Programmer')
		
ws.row('Wolfgang Mozart', 1756, 1791, 'Composer')
		
ws.row {
	ws.cell('Sigmund Freud')
	ws.cells(1856, 1939)
	ws.cell('Psychoanalyst')  
}

# we can do this manually,  but it will happen explicity
# due to write call,  which closes all open objects
# ws.table!

# this also applies to any open rows,  so this would be
# called automatically if necessary
# ws.row!

ss.write('test2.xlsx')

##############################################################################
__END__
