option explicit

'[path=\Element]
'[group=Element]

!INC Logging.LogManager

dim KEYWORD_TO_APPEND
' CHANGE THIS VALUE BEFORE RUNNING
KEYWORD_TO_APPEND = "infrastructure_view"

sub main
	dim logger
	set logger = LogManager.getLogger("Element.Append to Keyword")
	
	dim diagram as EA.Diagram
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element

	logger.INFO "Start..."

	'get the current diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		for each diagramObject in diagram.SelectedObjects		
			set element = Repository.GetElementByID(diagramObject.ElementID)
			logger.info "Working on '" & element.name & "'"
			if InStr(element.Tag, KEYWORD_TO_APPEND) = 0 then
				If element.Tag = "" then
					element.Tag = KEYWORD_TO_APPEND
				else
					element.Tag = element.Tag & "," & KEYWORD_TO_APPEND
				end if
				
				if not element.Update() then
					logger.WARN "Update failed: " & element.GetLastError()
					Session.Output
				end if
			end if
		next
	end if
	logger.INFO "Done"
end sub

main