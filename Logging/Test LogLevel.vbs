option explicit

'[path=\Logging]
'[group=Logging]

!INC Assert.Assert
!INC Logging.LogLevel

sub TestLogLevelNames
	assertEquals "Level_OFF has name", "OFF", Level_OFF.Name
	assertEquals "Level_FATAL has name", "FATAL", Level_FATAL.Name
	assertEquals "Level_ERROR has name", "ERROR", Level_ERROR.Name
	assertEquals "Level_WARN has name", "WARN", Level_WARN.Name
	assertEquals "Level_INFO has name", "INFO", Level_INFO.Name
	assertEquals "Level_DEBUG has name", "DEBUG", Level_DEBUG.Name
	assertEquals "Level_TRACE has name", "TRACE", Level_TRACE.Name
	assertEquals "Level_ALL has name", "ALL", Level_ALL.Name
end sub

sub TestIsMoreSpecificThan
	assertTrue "Level_OFF is more specific than Level_FATAL", Level_OFF.isMoreSpecificThan(Level_FATAL)
	assertTrue "Level_FATAL is more specific than Level_ERROR", Level_FATAL.isMoreSpecificThan(Level_ERROR)
	assertTrue "Level_ERROR is more specific than Level_WARN", Level_ERROR.isMoreSpecificThan(Level_WARN)
	assertTrue "Level_WARN is more specific than Level_INFO", Level_WARN.isMoreSpecificThan(Level_INFO)
	assertTrue "Level_INFO is more specific than Level_DEBUG", Level_INFO.isMoreSpecificThan(Level_DEBUG)
	assertTrue "Level_DEBUG is more specific than Level_TRACE", Level_DEBUG.isMoreSpecificThan(Level_TRACE)
	assertTrue "Level_TRACE is more specific than Level_ALL", Level_TRACE.isMoreSpecificThan(Level_ALL)
end sub

sub main
	TestLogLevelNames
	TestIsMoreSpecificThan
end sub

main