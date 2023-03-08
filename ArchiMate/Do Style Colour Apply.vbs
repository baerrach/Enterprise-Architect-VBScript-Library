option explicit

'[path=\ArchiMate]
'[group=ArchiMate]

!INC ArchiMate.Style Colour Apply
!INC Logging.LogManager

sub main
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Do Style Colour Apply")
	
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
		for each diagramObject in diagram.DiagramObjects		
			set element = Repository.GetElementByID(diagramObject.ElementID)
			set myArchiMateElement = new ArchiMateElement
			myArchiMateElement.init diagramObject, element
			logger.info "Working on '" & element.name & "'"
			applyStyleColour myArchiMateElement
		next
	end if
	logger.INFO "Done"
end sub

main