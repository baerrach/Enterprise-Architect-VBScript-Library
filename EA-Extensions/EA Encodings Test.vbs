option explicit

'[path=\EA-Extensions]
'[group=EA-Extensions]

!INC Assert.Assert
!INC Logging.LogManager
!INC EA-Extensions.EA Encodings

dim logger
set logger = LogManager.getLogger("EA-Extensions.EA Encodings")
LogManager.RootLogger.LogLevel = Level_INFO

sub TestToStringWhenNoEntriesEncodesAsEmpty
	dim OUT
	set OUT = new SparxKeyValueEncodedString
	assertEquals "Empty SparxKeyValueEncodedString encodes as empty string", "", OUT.toString()
end sub

sub TestFromStringWhenEmptyHasNoEntries
	dim OUT
	set OUT = new SparxKeyValueEncodedString
	assertEquals "Empty SparxKeyValueEncodedString encodes as empty string", 0, OUT.Count
end sub

sub TestToString
	dim OUT
	set OUT = new SparxKeyValueEncodedString
	OUT.Add "SX", "0"
	OUT.Add "SY", "0"
	OUT.Add "EX", "0"
	OUT.Add "EY", "0"
	OUT.Add "EDGE", "2"
	OUT.Add "$LLB",""
	OUT.Add "LLT", ""
	OUT.Add "LMT", ""
	OUT.Add "LMB", "CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0"
	OUT.Add "LRT",""
	OUT.Add "LRB", ""
	OUT.Add "IRHS", ""
	OUT.Add "ILHS", ""
	
	' Adding the same key shouldn't change the order, or add duplicate keys
	OUT.Add "SX", "0"
	
	assertEquals "toString()", "SX=0;SY=0;EX=0;EY=0;EDGE=2;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;", OUT.toString()

end sub

sub TestFromString
	dim OUT
	set OUT = new SparxKeyValueEncodedString
	OUT.fromString("SX=0;SY=0;EX=0;EY=0;EDGE=2;$LLB=;LLT=;LMT=;LMB=CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0;LRT=;LRB=;IRHS=;ILHS=;")
	assertEquals "SX=0", "0", OUT.Item("SX")
	assertEquals "SY=0", "0", OUT.Item("SY")
	assertEquals "EX=0", "0", OUT.Item("EX")
	assertEquals "EY=0", "0", OUT.Item("EY")
	assertEquals "EDGE=2", "2", OUT.Item("EDGE")
	assertEquals "$LLB=","", OUT.Item("$LLB")
	assertEquals "LLT=", "", OUT.Item("LLT")
	assertEquals "LMT=", "", OUT.Item("LMT")
	assertEquals "LMB=CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0", "CX=102:CY=13:OX=0:OY=0:HDN=0:BLD=0:ITA=0:UND=0:CLR=-1:ALN=1:DIR=0:ROT=0", OUT.Item("LMB")
	assertEquals "LRT=","", OUT.Item("LRT")
	assertEquals "LRB=", "", OUT.Item("LRB")
	assertEquals "IRHS=", "", OUT.Item("IRHS")
	assertEquals "ILHS=", "", OUT.Item("ILHS")
end sub

sub main
	TestToStringWhenNoEntriesEncodesAsEmpty
	TestFromStringWhenEmptyHasNoEntries
	TestFromString
	TestToString
	logger.INFO "Tests Completed"
end sub

main