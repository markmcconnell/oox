class Oox::Spreadsheet
	require 'oox/unique'
	require 'oox/spreadsheet/worksheet'
	
	######################################################################
	def self.write(file, style={}, &block)
		self.new(style, &block).write(file)
	end

	######################################################################
	attr_reader	:strings
	attr_accessor	:tableid

	# create a new spreadsheet object.  the +style+ parameter specifies
	# a default style for all created worksheets.  the attributes for this
	# parameter can be seen at Oox::Spreadsheet::Worksheet#style
	def initialize(style={})
		@strings = Oox::Unique.new
		@formats = Oox::Unique.new
		@fonts   = Oox::Unique.new
		@fills   = Oox::Unique.new
		@styles  = Oox::Unique.new

		@sheets  = []
		@tableid = 0
		
		## add a generic number format to the document
		format?(:id => '164', :code => '0.00')
		
		## add a default style
		style?(style)
		
		## excel seems to need a bogus fill in the second slot
		fill?(:bg => '010101') # whatever...
		
		## allow block style
		yield(self) if (block_given?)
	end
	
	######################################################################
	# :nodoc:
	def string?(v)
		@strings.id(v.to_s)
	end
	
	# get a numeric identifier for a specified format id and code
	def format?(v)
		id = v.fetch(:formatId, 0)
		
		@formats.id({ 
			:formatId   => id,
			:formatCode => v.fetch(:formatCode, nil) 
		}, id)
	end
	
	# get a numeric identifier for a specified font
	def font?(v)
		@fonts.id({ 
			:font   => v.fetch(:font, 'Calibri'),
			:size   => v.fetch(:size, '11'),
			:bold   => !!v.fetch(:bold,   false),
			:italic => !!v.fetch(:italic, false),
			:under  => !!v.fetch(:under,  false),
			:fg     => v.fetch(:fg, nil)
		})
	end
	
	# get a numeric identifier for a specified background fill
	def fill?(v)
		@fills.id({ 
			:bg 	=> v.fetch(:bg, nil) 
		})
	end
	
	# get a numeric identifier for a master style object
	def style?(v)
		@styles.id({
			:font   => font?(v),
			:fill   => fill?(v),
			:format => format?(v),
			:valign => v.fetch(:valign, :bottom),
			:halign => v.fetch(:halign, :left)
		})
	end
	
	######################################################################
	# create a new worksheet with the specified +name+ and add it to the
	# current spreadsheet.  most spreadsheet functions are accomplished
	# through the Oox::Spreadsheet::Worksheet object.
	def worksheet(name="sheet#{@sheets.length+1}")
		ws = Oox::Spreadsheet::Worksheet.new(self, name)
		
		@sheets.push(ws)
		yield(ws) if (block_given?)
		
		return(ws)
	end
	
	# write the current spreadsheet and all worksheets to the specified
	# file as a zipped openxml document archive.  this method is 
	# compatible with excel and other openxml document readers.
	def write(file, opts=Zip::ZipFile::CREATE)
		Zip::ZipFile.open(file, opts) {|zip|
			zip.dir.mkdir('_rels')
			zip.dir.mkdir('docProps')
			zip.dir.mkdir('xl')
			zip.dir.mkdir('xl/worksheets')
			zip.dir.mkdir('xl/worksheets/_rels')
			zip.dir.mkdir('xl/tables')
			zip.dir.mkdir('xl/_rels')
			
			@zip = zip
			write!
		}
		
		@zip = nil
		return(self)
	end
	
	######################################################################
	protected
	
	# :nodoc:
	# some internal methods for creating all the crazy xml necessary to
	# export our document.  the "rel" system for this format is wacky!
	def write_xml(file)
		xb = Builder::XmlMarkup.new
		xb.instruct!(:xml, :version => "1.0", :encoding => "UTF-8", :standalone => "yes")
		
		yield(xb)
		@zip.file.open(file, 'w') {|f| f.write(xb.target!) }
	end
	
	# :nodoc:
	# write out the enrtire file structure of the openxml zip file
	def write!
		write_xml('[Content_Types].xml') {|xb|
			xb.Types(:xmlns => "http://schemas.openxmlformats.org/package/2006/content-types") {|t|   
				t.Default(:Extension => "rels", :ContentType => "application/vnd.openxmlformats-package.relationships+xml")   
				t.Default(:Extension => "xml",  :ContentType => "application/xml")   
				
				t.Override(:PartName => "/xl/workbook.xml",   :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")   
				t.Override(:PartName => "/docProps/app.xml",  :ContentType => "application/vnd.openxmlformats-officedocument.extended-properties+xml")   
				t.Override(:PartName => "/docProps/core.xml", :ContentType => "application/vnd.openxmlformats-package.core-properties+xml")   
				
				@sheets.each_with_index {|n,idx|
					t.Override(:PartName => "/xl/worksheets/sheet#{idx+1}.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
					
					n.tables.each {|table|
						t.Override(:PartName => "/xl/tables/table#{table[:id]}.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
					}
				}
				
				t.Override(:PartName => "/xl/styles.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
				t.Override(:PartName => "/xl/sharedstrings.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") 
			}
		}
		
		write_xml('_rels/.rels') {|xb| 
			xb.Relationships(:xmlns => "http://schemas.openxmlformats.org/package/2006/relationships") {|r|
				r.Relationship(:Id => "rId1", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", :Target => "/xl/workbook.xml")
				r.Relationship(:Id => "rId2", :Type => "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", :Target => "/docProps/core.xml")
				r.Relationship(:Id => "rId3", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", :Target => "/docProps/app.xml")
			}
		}
		
		write_xml('docProps/core.xml') {|xb|
			xb.tag!("cp:coreProperties", "\n", "xmlns:cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "xmlns:dc" => "http://purl.org/dc/elements/1.1/", "xmlns:dcmitype"=> "http://purl.org/dc/dcmitype/", "xmlns:dcterms"=> "http://purl.org/dc/terms/", "xmlns:xsi"=> "http://www.w3.org/2001/XMLSchema-instance")
		}

		write_xml('docProps/app.xml') {|xb|
			xb.Properties(:xmlns => "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "xmlns:vt" => "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes") {|p|
				p.HeadingPairs {|h|
					h.tag!("vt:vector", :size => "2", :baseType => "variant") {|v|
						v.tag!("vt:variant") {|va| va.tag!("vt:lpstr", "Worksheets")   }
						v.tag!("vt:variant") {|va| va.tag!("vt:i4",    @sheets.length) }
					}
				}
			}
		}
		
		write_xml('xl/sharedstrings.xml') {|xb|
			xb.sst(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", :count => @strings.length, :uniqueCount => @strings.length) {|s|
				@strings.to_a.each {|string|
					# builder is too slow for this,  so we just encode
					# add the string directly;  this is a little 
					# hacky,  but it's pretty fast.
					s << "<si><t>#{string.to_xs}</t></si>"
				}
			}
		}
		
		write_xml('xl/styles.xml') {|xb|
			xb.styleSheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") {|s|
				fmts = @formats.keys.select {|n| n[:formatCode] }
				s.numFmts(:count => fmts.length.to_s) {|f|
					fmts.each {|fmt|
						f.numFmt(:numFmtId => fmt[:formatId], :formatCode => fmt[:formatCode])
					}
				}
				
				s.fonts(:count => @fonts.length.to_s) {|f|
					@fonts.to_a.each {|t|
						f.font {|p|
							p.name(:val  => t[:font])
							p.sz(:val    => t[:size])
							p.color(:rgb => t[:fg]) if (t[:fg])
							p.b if (t[:bold])
							p.i if (t[:italic])
							p.u if (t[:under])
						}
					}
				}
				
				s.fills(:count => @fills.length.to_s) {|f|
					@fills.to_a.each {|t|
						f.fill {|p|
							if (t[:bg])
								p.patternFill(:patternType => 'solid') {|pt|
									pt.fgColor(:rgb => "FF#{t[:bg]}")
									pt.bgColor(:indexed => "64")
								}
							else
								p.patternFill(:patternType => 'none')
							end
						}
					}
				}
				
				s.borders(:count => "1") {|f| f.border }
				s.cellStyleXfs(:count => "1") {|f| f.xf }
				
				s.cellXfs(:count => @styles.length.to_s) {|f|
					@styles.to_a.each {|t|
						o = {
							:numFmtId => t[:format].to_s, 
							:fontId   => t[:font].to_s, 
							:fillId   => t[:fill].to_s, 
							:borderId => "0", 
							:xfId     => "0"
						}
						o[:applyFill] = "1" unless (t[:fill].zero?)
						o[:applyFont] = "1" unless (t[:font].zero?)
						
						if ((t[:valign] != :bottom) || (t[:halign] != :left))
							f.xf(o) {|x| x.alignment(:vertical => t[:valign].to_s, :horizontal => t[:halign].to_s) }
						else
							f.xf(o)
						end
					}
				}
			}
		}

		write_xml('xl/workbook.xml') {|xb|
			xb.workbook(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r"=> "http://schemas.openxmlformats.org/officeDocument/2006/relationships") {|w|
				w.workbookPr(:date1904 => "0")
				
				w.bookViews {|v|
					v.workbookView(:xWindow => "0", :yWindow => "0", :windowWidth => "22667", :windowHeight => "17000", :tabRatio => "500")
				}
				
				w.sheets {|s|
					@sheets.each_with_index {|n,idx|
						s.sheet(:name => n.name, :sheetId => (idx+1).to_s, "r:id" => "rId#{(idx+2)}")
					}
				}
				
				w.definedNames("\n")
			}
		}
		
		write_xml('xl/_rels/workbook.xml.rels') {|xb|
			xb.Relationships(:xmlns => "http://schemas.openxmlformats.org/package/2006/relationships") {|r|
				r.Relationship(:Id => "rId0", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", :Target => "sharedstrings.xml")
				r.Relationship(:Id => "rId1", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", :Target => "/xl/styles.xml")
				@sheets.length.times {|n|
					r.Relationship(:Id => "rId#{n+2}", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", :Target => "/xl/worksheets/sheet#{n+1}.xml")
				}
			}
		}
		
		@sheets.each_with_index {|sheet,sidx| 
			write_xml("xl/worksheets/sheet#{sidx+1}.xml") {|xb|
				xb.worksheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships") {|w|
					w.sheetViews {|v|
						v.sheetView("\n", :workbookViewId => "0")
					}
					
					w.cols {|c|
						sheet.width.each_with_index {|width,idx|
							idx = (idx+1).to_s
							c.col(:min => idx, :max => idx, :width => width.to_f.to_s, :bestFit => "1", :customWidth => "1")
						}
					}
					
					# builder is too slow to do this programatically,  we 
					# just add the raw xml as a string instead (way faster!)
					w << sheet.to_s
					
					w.pageMargins(:left => "0.7", :right => "0.7", :top => "0.75", :bottom => "0.75", :header => "0.3", :footer => "0.3")
					
					if (sheet.tables.any?)
						w.tableParts(:count => sheet.tables.length.to_s) {|t|
							sheet.tables.length.times {|n|
								t.tablePart("r:id" => "rId#{n+1}")
							}
						}
					end
				}
				
			}
			next unless (sheet.tables.any?)
			
			write_xml("xl/worksheets/_rels/sheet#{sidx+1}.xml.rels") {|xb|
				xb.Relationships(:xmlns => "http://schemas.openxmlformats.org/package/2006/relationships") {|r|
					sheet.tables.each_with_index {|table,idx|
						r.Relationship(:Id => "rId#{idx+1}", :Target => "../tables/table#{table[:id]}.xml", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table") 
					}
				}
			}
			
			sheet.tables.each {|table|
				write_xml("xl/tables/table#{table[:id]}.xml") {|xb|
					xb.table(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", :id => table[:id], :name => table[:name], :displayName => table[:name], :ref => table[:range], :totalsRowShown => "0") {|t|
						t.autoFilter(:ref => table[:range])
						
						t.tableColumns(:count => table[:header].length) {|c|
							table[:header].each_with_index {|str,idx|
								c.tableColumn(:id => "#{idx+1}", :name => str)
							}
						}
						
						t.tableStyleInfo(table[:style])
					}
				}
			}
		}
	end
end

##############################################################################
__END__
