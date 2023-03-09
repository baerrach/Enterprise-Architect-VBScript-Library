'[path=\Logging]
'[group=Logging]

!INC Local Scripts.EAConstants-VBScript
!INC Logging.LogLevel

const LOG_TAB_NAME = "Logging"

Class LoggerClass

	private m_Name
	private m_LogLevel

	Private Sub Class_Initialize
		m_Name = "<unnamed>"
		set m_LogLevel = Level_ALL
	End Sub
	
	Public Sub init(name)
		m_Name = name
	End Sub
	
	' name property
	Public Property Get Name
		Name = m_Name
	End Property
	
	' logLevel property
	Public Property Get LogLevel
		set LogLevel = m_LogLevel
	End Property
	Public Property Let LogLevel(value)
		set m_LogLevel = value
	End Property
	
	public function isEnabled(msgLevel)
		isEnabled = msgLevel.isMoreSpecificThan(m_LogLevel)
	end function

	Public sub log(msgLevel, message)
		if isEnabled(msgLevel) then
			Repository.WriteOutput LOG_TAB_NAME, now() & " " & msgLevel.Name & " [" & m_Name & "] " & message, 0
		end if
	end sub

	Public sub trace(message)
		log Level_TRACE, message
	end sub
	
	Public sub debug(message)
		log Level_DEBUG, message
	end sub
	
	Public sub info(message)
		log Level_INFO, message
	end sub
	
	Public sub warn(message)
		log Level_WARN, message
	end sub
	
	Public sub error(message)
		log Level_ERROR, message
	end sub
end Class


