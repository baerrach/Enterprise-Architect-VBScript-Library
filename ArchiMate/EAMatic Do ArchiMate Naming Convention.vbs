option explicit

'[path=\ArchiMate]
'[group=ArchiMate]
'


!INC Local Scripts.EAConstants-VBScript
!INC Utils.Util
!INC Logging.LogManager
!INC ArchiMate.Style Colour Apply

function EA_GetMenuItems(MenuLocation, MenuName)
'	if MenuLocation <> "Diagram" then
'		MenuLocation = Empty
'		exit function
'	end if
	
    if MenuName = "" then
        'Menu Header
        EA_GetMenuItems = "-&ArchiMate"
    else
        if MenuName = "-&ArchiMate" then
            'Menu items
            Dim menuItems(1)
            menuItems(0) = "Do ArchiMate Naming Convention"
            EA_GetMenuItems = menuItems
         end if
    end if
end function

'react to user clicking a menu option
function EA_MenuClick(MenuLocation, MenuName, ItemName)
	if MenuName = "-&ArchiMate" then
        Select Case ItemName
            case "Do ArchiMate Naming Convention"
                handleDoArchiMateNamingConvention
        end select
    end if
end function

function handleDoArchiMateNamingConvention
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
end function