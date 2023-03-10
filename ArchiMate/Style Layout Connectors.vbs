'[path=\ArchiMate]
'[group=ArchiMate]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC EA-Extensions.All
!INC Logging.LogManager
!INC ArchiMate.ArchiMateElement

sub applyStyleLayoutConnectors(myArchiMateElement)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Style Layout Connectors")

	if myArchiMateElement.stereotype = "" then
		logger.Info "Ignoring non-ArchiMate element name=" & myArchiMateElement.element.name & " stereotype=" & myArchiMateElement.element.stereotype & " type=" & myArchiMateElement.element.type
		exit sub
	end if
	
	logger.debug "ElementID=" & myArchiMateElement.Element.ElementID
	logger.debug "Diagram's diagramID=" & myArchiMateElement.DiagramObject.DiagramID

	dim diagram, diagramEx
	set diagram = Repository.GetDiagramByID(myArchiMateElement.DiagramObject.DiagramID)
	set diagramEx = DiagramExtension.createForDiagram(diagram)

	Dim connector, connectors, diagramLink
	set connectors = myArchiMateElement.element.Connectors
	
	Dim filteredConnectors()
	Redim filteredConnectors(connectors.Count)
	Dim connectorEnd
	for each connector in connectors
		logger.debug "Connector.ConnectorID=" & connector.ConnectorID

		logger.debug "  DiagramID=" & connector.DiagramID
		logger.debug "  SupplierID=" & connector.SupplierID
		logger.debug "  ClientID=" & connector.ClientID


		logger.debug "  StartPointX=" & connector.StartPointX
		logger.debug "  StartPointY=" & connector.StartPointY

		logger.debug "  EndPointX=" & connector.EndPointX
		logger.debug "  EndPointY=" & connector.EndPointY

		logger.debug "  Type=" & connector.Type
		logger.debug "  MetaType=" & connector.MetaType
		logger.debug "  RouteStyle=" & connector.RouteStyle
		logger.debug "  Stereotype=" & connector.Stereotype
		
		if connector.DiagramId <> 0 or true then
			set diagramLink = diagramEx.getDiagramLinkForConnector(connector)
			logger.debug "  diagramLink.ConnectorID=" & diagramLink.ConnectorID
			logger.debug "  diagramLink.Geometry=" & diagramLink.Geometry
			logger.debug "  diagramLink.LineStyle=" & diagramLink.LineStyle
			logger.debug "  diagramLink.Path=" & diagramLink.Path
			logger.debug "  diagramLink.SourceInstanceUID=" & diagramLink.SourceInstanceUID
			logger.debug "  diagramLink.TargetInstanceUID=" & diagramLink.TargetInstanceUID
		end if

	next	

end sub
