Gem::Specification.new do |s|
	s.name        = "oox"
	s.version     = "0.1"
	s.license     = "GPL-2"
	s.authors     = ["Mark Anthony McConnell"]
	s.email       = ["msg-ox@markmcconnell.us"]
	s.homepage    = "https://markmcconnell.us/oox"
	s.summary     = %{Writes Office OpenXML Spreadsheet and Document files.}
	s.description = %{Office OpenXML writer with good support for data tables, numeric formats and styles.}
	
	s.add_runtime_dependency 'builder'
	s.add_runtime_dependency 'rubyzip'
	
	s.files = [
		"lib/oox.rb", 
		"lib/oox/unique.rb",
		"lib/oox/spreadsheet.rb",
		"lib/oox/spreadsheet/worksheet.rb"
	]

	s.require_paths = ["lib"]
end
