option explicit

'[path=\Logging]
'[group=Logging]

!INC Assert.Assert
!INC Logging.LogManager

dim OUT
set OUT = new LogManagerClass

sub TestParentNameWhenEmptyString
	assertEquals "when name is empty string then parent is empty string", "", OUT.parentName("")
end sub

sub TestParentNameWhenNoHierarchy
	assertEquals "when name has not hierarchy then parent is empty string", "", OUT.parentName("no_hierarchy")
end sub

sub TestParentNameWhenOneLevelOfHierarchy
	assertEquals "when name has one level of hierarchy then parent", "one", OUT.parentName("one.level_of_hierarchy")
end sub

sub TestParentName
	TestParentNameWhenEmptyString
	TestParentNameWhenNoHierarchy
	TestParentNameWhenOneLevelOfHierarchy
end sub

sub TestLogManagerExists
	assertNotNothing "LogManager should be exist", LogManager
end sub

sub TestLogManagerGetLogger
	dim logger
	set logger = LogManager.getLogger("test logger")
	assertNotNothing "Logger should exist", logger
	assertEquals "logger name", "test logger", logger.name
end sub

sub TestLogManagerGetLoggerInheritsParentLevel
	dim child, parent

	set parent = LogManager.getLogger("parent")
	assertSame "parent inherits root logger log level", LogManager.RootLogger.LogLevel, parent.LogLevel
	parent.LogLevel = Level_DEBUG
	
	set child  = LogManager.getLogger("parent.child")
	assertSame "child inherits parent log level", parent.LogLevel, child.LogLevel
end sub

sub main
	TestParentName
	TestLogManagerExists
	TestLogManagerGetLogger
	TestLogManagerGetLoggerInheritsParentLevel
end sub

main