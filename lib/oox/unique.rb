class Oox::Unique < Hash
	def id(v,i=nil)
		return(self[v]) if (has_key?(v))
		self[v] = i.nil? ? self.length : i
	end
	
	def to_a
		super.sort! {|a,b| a.last <=> b.last }.collect! {|a| a.first }
	end
end

##############################################################################
__END__
