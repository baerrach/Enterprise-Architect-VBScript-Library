'[path=\Logging]
'[group=Logging]

class LogLevel
	private m_name
	private m_intLevel
	
	Private Sub Class_Initialize
		m_Name = "<unnamed>"
		m_intLevel = 0
	End Sub
	
	public sub init(name, intLevel)
		m_Name = name
		m_intLevel = intLevel
	end sub
	
	' name property
	Public Property Get Name
		Name = m_Name
	End Property
	
	' intLevel property
	Public Property Get intLevel
		intLevel = m_intLevel
	End Property

	Public function isMoreSpecificThan(otherLevel)
		isMoreSpecificThan = intLevel <= otherLevel.intLevel
	end function
	
end class

dim Level_OFF, Level_FATAL, Level_ERROR, Level_WARN, Level_INFO, Level_DEBUG, Level_TRACE, Level_ALL
set Level_OFF = new LogLevel
set Level_FATAL = new LogLevel
set Level_ERROR = new LogLevel
set Level_WARN = new LogLevel
set Level_INFO = new LogLevel
set Level_DEBUG = new LogLevel
set Level_TRACE = new LogLevel
set Level_ALL = new LogLevel

Level_OFF.init "OFF", 		0
Level_FATAL.init "FATAL", 100
Level_ERROR.init "ERROR", 200
Level_WARN.init "WARN",   300
Level_INFO.init "INFO",   400
Level_DEBUG.init "DEBUG", 500
Level_TRACE.init "TRACE", 600
Level_ALL.init "ALL",   32767