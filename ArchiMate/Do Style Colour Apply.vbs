'[path=\ArchiMate]
'[group=ArchiMate]

!INC ArchiMate.Style Colour Apply

sub main
	dim diagram as EA.Diagram
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	dim myArchiMateElement

	'get the current diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		for each diagramObject in diagram.DiagramObjects		
			set element = Repository.GetElementByID(diagramObject.ElementID)
			set myArchiMateElement = new ArchiMateElement
			myArchiMateElement.init diagramObject, element
			applyStyleColour myArchiMateElement
		next
	end if
	Session.Output "done"
end sub

main