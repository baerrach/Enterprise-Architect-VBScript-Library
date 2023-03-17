option explicit

'[path=\ArchiMate]
'[group=ArchiMate]

!INC ArchiMate.Style Access Relation Unspecified Navigability Apply
!INC Logging.LogManager

sub main
	dim logger
	set logger = LogManager.getLogger("Do Style Access Relation Unspecified Navigability Apply")
	
	dim diagram as EA.Diagram
	dim connector as EA.Connector
	dim connectorEd as EA.ConnectorEnd

	logger.INFO "Start..."

	'get the current diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		logger.debug "diagram.DiagramID=" & diagram.DiagramID
		set connector = diagram.SelectedConnector
		if connector is Nothing then
			logger.warn "No connector selected"
			exit sub
		end if

		logger.info "Working on connector.ConnectorID=" & connector.ConnectorID
		if applyStyleAccessRelationUnspecifiedNavigability(connector) then
			' Connector.Update() and ConnectorEnd.Update() don't appear to update the Diagram, force a reload
			Repository.ReloadDiagram diagram.DiagramID
		end if
	end if
	logger.INFO "Done"
end sub

main
