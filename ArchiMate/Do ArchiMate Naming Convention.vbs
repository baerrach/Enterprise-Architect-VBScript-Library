option explicit

'[path=\ArchiMate]
'[group=ArchiMate]

!INC ArchiMate.Naming Convention
!INC Logging.LogManager

'
' On the selected objects in the current diagram, apply ArchiMate Naming Convention
'
sub main
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Do ArchiMate Naming Convention")
	
	dim diagram as EA.Diagram
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	dim myArchiMateElement

	logger.INFO "Start..."

	'get the current diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		for each diagramObject in diagram.SelectedObjects		
			set element = Repository.GetElementByID(diagramObject.ElementID)
			logger.info "Working on '" & element.name & "'"
			applyArchiMateNamingConventionToElement element
		next
	end if
	logger.INFO "Done"
end sub

main