option explicit

'[path=\ArchiMate]
'[group=ArchiMate]
'
'EA-Matic
'EA-Matic: http://bellekens.com/ea-matic/

!INC Local Scripts.EAConstants-VBScript
!INC Utils.Util
!INC Logging.LogManager
!INC ArchiMate.Style Colour Apply
!INC ArchiMate.Style Size Apply
!INC ArchiMate.Naming Convention

function EA_OnPostNewDiagramObject(Info)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.EA-Matic Apply ArchiMate Conventions On New Diagram Object")
	
	logger.INFO "Start..."

	'''gather required details
	dim elementId
	elementId = Info.Get("ID")
	logger.DEBUG " ID=" & elementId

	dim diagramID
	diagramID = Info.Get("DiagramID")
	logger.DEBUG " DiagramID=" & diagramID

	dim DUID
	DUID = Info.Get("DUID")
	logger.DEBUG " DUID=" & DUID

	dim diagram as EA.Diagram
	dim diagramObject as EA.DiagramObject
	dim element as EA.Element
	dim myArchiMateElement

	set diagram = Repository.getDiagramByID(diagramID)
	set element = Repository.GetElementByID(elementId)
	set diagramObject = getDiagramObjectByElementId(diagram, element.elementId)

	'''validate details and provide error message

	'''Delegate actual work to included script
	set myArchiMateElement = new ArchiMateElement
	myArchiMateElement.init diagramObject, element
	applyStyleColour myArchiMateElement
	applyStyleSize myArchiMateElement
	applyArchiMateNamingConventionToElement element

	logger.INFO "Done"
end function