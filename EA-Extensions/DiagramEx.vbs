'[path=\EA-Extensions]
'[group=EA-Extensions]

!INC Logging.LogManager

class DiagramExtenionClass
	private logger
	private m_Diagram

	Private Sub Class_Initialize
		set logger = LogManager.getLogger("DiagramEx")
		set m_Diagram = nothing
	End Sub
	
	Public sub init(diagram)
		set m_diagram = diagram
	end sub
	
	function getDiagramLinkForConnector(connector)	
		dim diagramLink, diagramLinks
		
		logger.debug "Looking for matching diagramLink.ConnectorID=" & connector.ConnectorID
		logger.debug "  connector.DiagramId=" & connector.DiagramId

		set diagramLinks = m_diagram.DiagramLinks
		for each diagramLink in diagramLinks
			logger.debug "diagramLink.ConnectorID=" & diagramLink.ConnectorID
			if diagramLink.ConnectorID = connector.ConnectorID then
				set getDiagramLinkForConnector = diagramLink
				exit function
			end if
		next
		
		Err.Raise vbObjectError + 1, "getDiagramLinksForConnector", "No diagramLink matches connector with ConnectorID=" & connector.ConnectorID
	end function
		
	'returns the diagram object by the specified element ID from the given diagram
	function getDiagramObjectByElementId(elementId)
		set getDiagramObjectByElementId = nothing

		dim diagramObject as EA.DiagramObject
		dim element as EA.Element
		for each diagramObject in m_diagram.DiagramObjects
			if diagramObject.ElementID = CLng(elementId) then
				set getDiagramObjectByElementId = diagramObject
				exit for
			end if
		next
	end function
end class

class DiagramExtensionNamespace
	private logger

	Private Sub Class_Initialize
		set logger = LogManager.getLogger("DiagramEx")
	End Sub
	
	function createForDiagram(diagram)
		dim result
		set result = new DiagramExtenionClass
		result.init diagram
		set createForDiagram = result
	end function
end class

dim DiagramExtension
set DiagramExtension = new DiagramExtensionNamespace
