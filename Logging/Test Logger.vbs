option explicit

'[path=\Logging]
'[group=Logging]

!INC Assert.Assert
!INC Logging.Logger
!INC Logging.LogManager

sub TestLoggerName
	dim logger
	set logger = new LoggerClass
	logger.init "test logger"
	assertEquals "logger name", "test logger", logger.name
end sub

sub TestLoggerIsEnabledDefaultsToLevelAll
	dim logger
	set logger = new LoggerClass
	logger.init "test logger"
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_OFF)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_FATAL)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_WARN)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_INFO)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_DEBUG)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_TRACE)
	assertTrue "A logger without a level set is always enabled", logger.isEnabled(Level_ALL)
end sub

sub TestLoggerIsEnabled
	TestLoggerIsEnabledDefaultsToLevelAll
end sub

sub TestLoggerDebugWhenEnabled
	dim logger
	set logger = new LoggerClass
	logger.init "test logger"
	logger.LogLevel = Level_DEBUG
	logger.debug "TestLoggerDebugWhenEnabled: this message is logged"
end sub

sub TestLoggerDebugWhenDisabled
	dim logger
	set logger = new LoggerClass
	logger.LogLevel = Level_INFO
	logger.init "test logger"
	logger.debug "TestLoggerDebugWhenDisabled: this message wont be logged"
end sub

sub main
	LogManager.clear

	TestLoggerName
	TestLoggerIsEnabled
	TestLoggerDebugWhenEnabled
	TestLoggerDebugWhenDisabled
end sub

main