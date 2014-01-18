##############################################################################
## we build raw XML because it is faster,  and because it is safe
## to do so;  we are only creating tags and attributes with numeric
## data -- it is implicitly safe to save in an XML file.  
class Oox::Spreadsheet::Worksheet < String
	COLNAMES   = [('A'..'Z').to_a, ('AA'..'ZZ').to_a, ('AAA'..'ZZZ').to_a].flatten.freeze
	CELL_NUM   = ''.freeze
	CELL_STR   = ' t="s"'.freeze

	# :nodoc:
	attr_reader :id, :name, :tables
	def initialize(parent, name)
		@name   = name
		@parent = parent
		
		@rows   = 0
		@autow  = true
		@width  = []
		@tables = []
		
		@ropen  = false
		@topen  = false
		
		@styleh = {}
		@style  = 0
		
		self << %Q|<sheetData>|
	end
	
	######################################################################
	# merge the supplied attributes in +h+ into the default style for
	# the current worksheet, if run with a block,  then the previous
	# style is restored upon return.  available attributes are:
	#	:font		=> 'Calibri'		# font face
	#	:size		=> '11',		# font size
	#	:bold		=> true,		# bold font?
	#	:italic		=> true,		# italic font?
	#	:under		=> true,		# underline font?
	#	:fg		=> 'AABBCC'		# hex foreground color
	#	:bg		=> 'AABBCC'		# hex background color
	#	:formatId	=> 22,			# numeric format code
	#	:formatCode	=> '#,##0.00'		# format code data
	#	:valign		=> :middle,		# vertical alignment
	#	:halign		=> :center		# horizontal alignment
	def style(h={})
		last_s = @style
		last_h = @styleh.clone
		
		@style = @parent.style?(@styleh.merge!(h))
		if (block_given?)
			yield(self)
			
			@style  = last_s
			@styleh = last_h
		end
		return(self)
	end
	
	# set only the supplied attributes in +h+ as the default style for
	# the current worksheet,  if run with a block,  then the previous
	# style is restored upon return.  the attributes are the same as #style.
	def style!(h={})
		last_s = @style
		last_h = @styleh
		
		@style = @parent.style?(@styleh = h)
		if (block_given?)
			yield(self)
			
			@style  = last_s
			@styleh = last_h
		end
		return(self)
	end
	
	######################################################################
	# set the default widths of the worksheet columns to the values in +w+,
	# after these values are set,  any automatic width calculation is 
	# canceled and no longer performed.
	def width(*w)
		return(@width) if (w.empty?)
		
		@autow = false
		@width = w.clone
		return(self)
	end
	
	######################################################################
	# closes the current table header,  this sets the width of the current
	# table and stops any further cells in this row from being part of the
	# header.  this will be called automatically when starting a new row,
	# so it is only necessary to call this method when you want to add
	# datta to the "right" of your table on the current worksheet.
	def header!
		# autocall: worksheet close, table close, row close
		if (@hopen)
			@tables.last[:range] << COLNAMES[@cells]
			@hopen = false
		end
		return(self)
	end
	
	# closes the current table,  this sets the height of the current
	# table and stops any further rows from being a part of the table.
	# this will be called automatically when closing the worksheet,  so it
	# is only necessary to call this method when you want to add data
	# "below" your table on the current worksheet.
	def table!
		# autocall: table start, worksheet close
		header!
		
		if (@topen)
			@tables.last[:range] << @rows.to_s
			@topen = false
		end
		return(self)
	end
	
	# start a new table object with attributes specified by +style+,
	# available attributes are:
	#	:name		 	=> "TableN"		# the name of the table
	#	:style		 	=> "TableStyleMedium9"	# the name of the table style
	#	:showFirstColumn	=> 0			# table style attribute
	#	:showLastColumn		=> 0			# table style attribute
	#	:showRowStripes		=> 1			# table style attribute
	#	:showColumnStripes	=> 0			# table style attribute
	def table(style={})
		# close any previous table
		table!
		
		# mark table and header open
		@hopen = true
		@topen = true
		
		# start table structure
		@tables.push({
			:id     => (@parent.tableid += 1).to_s,
			:name   => (style.delete(:name) || "Table#{@parent.tableid}"),
			:header => [],			# list of strings
			:range  => COLNAMES[@cells+1] + @rows.to_s + ':',
			:style  => {
				:name 		   => style.fetch(:style, 'TableStyleMedium9'),
				:showFirstColumn   => style.fetch(:showFirstColumn,   0).to_s,
				:showLastColumn    => style.fetch(:showLastColumn,    0).to_s,
				:showRowStripes    => style.fetch(:showRowStripes,    1).to_s,
				:showColumnStripes => style.fetch(:showColumnStripes, 0).to_s
			}
		})
		
		if (block_given?)
			yield(self)
			table!
		end
		return(self)
	end
	
	######################################################################
	# closes the current row.  this closes any open table header and 
	# finalizes the xml output for this row.  this will be called
	# automatically when starting a new row,  so it is unecessary to call
	# this method directly.
	def row!
		# autocall: spreadsheet close, row start
		return(self) unless (@ropen)
		
		# close any header when new row begins
		header!
		
		@ropen  = false
		self << %Q|</row>|
	end
	
	# start a new row object.  any data in +cells+ will be automatically
	# added to the newly created row.  any cells added through this method
	# will be created with the default worksheet style.
	def row(*cells)
		# start row xml while ensuring previous row is closed
		@rows  += 1
		row! << %Q|<row r="#{@rows}">|
		
		# mark row open,  reset cell counter
		@ropen = true
		@cells = -1
		
		# fill cells
		cells.each {|c| cell(c) }
		if (block_given?)
			yield(self)
			row!
		end
		
		return(self)
	end
	
	# add an array of arrays as rows of columns to the current worksheet
	# using the default style.  useful for moving structure data.
	def rows(*rows)
		rows.each {|r| row(*r) }
		return(self)
	end
	
	######################################################################
	# add a cell to the current worksheet row,  it is your responsibility
	# to ensure a row has been started.  the +style+ attribute can have
	# the same values as the #style method,  but will only be applied to
	# this single cell
	def cell(data, style=@style)
		# set @cells to the _target_ cell index
		@cells += 1
		
		# calculate length for auto-width data if no width set
		if (@autow && ((width = data.to_s.length) >= (@width[@cells] || 0)))
			@width[@cells] = width + 1
		end
		
		# add raw cell data to open table header
		@tables.last[:header].push(data) if (@hopen)
		
		# determine document style and data type
		style  = style.kind_of?(Hash) ? @parent.style?(@styleh.merge(style)) : style
		string = if (data.kind_of?(Numeric))
			CELL_NUM # ''
		else
			# index string into sharedstrings table
			data = @parent.string?(data)
			CELL_STR # ' t="s"'
		end
		
		# build column xml
		self << %Q|<c r="#{COLNAMES[@cells]}#{@rows}" s="#{style}"#{string}><v>#{data}</v></c>|
	end
	
	# add multiple cells with the default style to the current row
	def cells(*data)
		data.each {|d| cell(d) }
		return(self)
	end
	
	# close any open rows and tables,  finalize the worksheet xml data
	# and then return a string representation of the <sheetData> xml.
	# it is generally uncessary to call this method directly.
	def to_s
		row!	# ensure row is closed
		table!	# ensure table is closed
		
		self << %Q|</sheetData>|
		super()
	end
end

##############################################################################
__END__
