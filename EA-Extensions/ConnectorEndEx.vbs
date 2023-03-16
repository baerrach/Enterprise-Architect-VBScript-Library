'[path=\EA-Extensions]
'[group=EA-Extensions]

!INC Logging.LogManager

class ConnectorEndExtenionClass
	private logger
	dim m_ConnectorEnd 
	
	Private Sub Class_Initialize
		set logger = LogManager.getLogger("ConnectorEndExtenion")
		set m_ConnectorEnd = nothing
	End Sub
	
	Public sub init(connectorEnd)
		set m_ConnectorEnd = connectorEnd
	end sub
	
end class

class ConnectorEndExtensionNamespace
	private logger

	Private Sub Class_Initialize
		set logger = LogManager.getLogger("ConnectorEndEx")
	End Sub
	
	function createForConnectorEnd(connectorEnd)
		dim result
		set result = new ConnectorEndExtenionClass
		result.init connectorEnd
		set createForConnectorEnd = result
	end function
	
	Public Property Get Navigable_Navigable
		Navigable_Navigable = "Navigable"
	End Property
	Public Property Get Navigable_NonNavigable
		Navigable_NonNavigable = "Non-Navigable"
	End Property
	Public Property Get Navigable_Unspecified
		Navigable_Unspecified = "Unspecified"
	End Property	
end class

dim ConnectorEndExtension
set ConnectorEndExtension = new ConnectorEndExtensionNamespace
