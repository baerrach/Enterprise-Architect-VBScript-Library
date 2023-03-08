'[path=\Logging]
'[group=Logging]

!INC Logging.Logger

Class LogManagerClass

	private m_rootLogger
	private m_currentLoggers
	private m_loggerConfig
	
	Private Sub Class_Initialize	  
		set m_rootLogger = new Logger
		m_rootLogger.init "<root logger>"
		set m_currentLoggers = CreateObject("Scripting.Dictionary")
	End Sub
	
	Public Sub init(loggerConfig)
		set m_loggerConfig = loggerConfig
	End Sub
	
	' RootLogger property
	Public Property Get RootLogger
		set RootLogger = m_rootLogger
	End Property
	
	Public Function parentName(name)
		if name = "" then
			parentName = ""
			exit function
		end if
		
		dim nameParts, i
		nameParts = split(name, ".")
		
		if UBound(nameParts)-LBound(nameParts)+1 = 1 then
			parentName = ""
			exit function
		end if
		
		parentName = nameParts(0)
		for i = 1 to UBound(nameParts)-1
			parentName = parentName & "." & nameParts(i)
		next
		
	end function
	
	Public function getLogger(name)
		if name = "" then
			set getLogger = m_rootLogger
			exit function
		end if
		
		if not m_currentLoggers.exists(name) then
			dim newLogger, parentLogger
			set parentLogger = getLogger(parentName(name))
			set newLogger = new Logger
			newLogger.init(name)
			newLogger.LogLevel = parentLogger.LogLevel
			m_currentLoggers.Add name, newLogger
		end if

		set getLogger = m_currentLoggers(name)
	end function
	
	Public sub clear()
		Repository.CreateOutputTab LOG_TAB_NAME
		Repository.ClearOutput LOG_TAB_NAME
		Repository.EnsureOutputVisible LOG_TAB_NAME	
	end sub	
end Class

dim LogManager
set LogManager = new LogManagerClass
LogManager.clear