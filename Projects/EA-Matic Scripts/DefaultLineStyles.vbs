'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' Script Name: DefaultLineStyles bla
' Author: Geert Bellekens
' Purpose: Allows to determine the standard style of new connectors when they are created on a diagram
' Date: 14/03/2015
'
dim lsDirectMode, lsAutoRouteMode, lsCustomMode, lsTreeVerticalTree, lsTreeHorizontalTree, _
lsLateralHorizontalTree, lsLateralVerticalTree, lsOrthogonalSquareTree, lsOrthogonalRoundedTree

lsDirectMode = "1"
lsAutoRouteMode = "2" 
lsCustomMode = "3"
lsTreeVerticalTree = "V"
lsTreeHorizontalTree = "H"
lsLateralHorizontalTree = "LH"
lsLateralVerticalTree = "LC"
lsOrthogonalSquareTree = "OS"
lsOrthogonalRoundedTree = "OR"

dim defaultStyle
dim menuDefaultLines


'*********EDIT BETWEEN HERE*************
' set here the menu name
menuDefaultLines = "&Set default linestyles"

' set here the default style to be used
defaultStyle = lsOrthogonalSquareTree

' set there the style to be used for each type of connector
function determineStyle(connector)
	dim connectorType
	connectorType = connector.Type
	select case connectorType
		case "ControlFlow", "StateFlow","ObjectFlow","InformationFlow"
			determineStyle = lsOrthogonalRoundedTree
		case "Generalization", "Realization", "Realisation"
			determineStyle = lsTreeVerticalTree
		case "UseCase", "Dependency","NoteLink", "Abstraction"
			determineStyle = lsDirectMode
		case else
			determineStyle = defaultStyle
	end select
end function
'************AND HERE****************

'the event called by EA
function EA_OnPostNewConnector(Info)
	'get the connector id from the Info
	dim connectorID
	connectorID = Info.Get("ConnectorID")
	dim connector 
	set connector = Repository.GetConnectorByID(connectorID)
	'get the current diagram
	dim diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		'first save the diagram
		Repository.SaveDiagram diagram.DiagramID
		'get the diagramlink for the connector
		dim diagramLink
		set diagramLink = getdiagramLinkForConnector(connector, diagram)
		if not diagramLink is nothing then
			'set the connectorstyle
			setConnectorStyle diagramLink, determineStyle(connector)
			'save the diagramlink
			diagramLink.Update
			'reload the diagram to show the link style
			Repository.ReloadDiagram diagram.DiagramID
		end if
	end if
end function



'Tell EA what the menu options should be
function EA_GetMenuItems(MenuLocation, MenuName)
	if MenuName = "" and MenuLocation = "Diagram" then
		'Menu Header
		EA_GetMenuItems = menuDefaultLines
	end if 
end function


'react to user clicking a menu option
function EA_MenuClick(MenuLocation, MenuName, ItemName)
	if ItemName = menuDefaultLines then
		dim diagram 
		dim diagramLink
		dim connector
		dim dirty
		dirty = false
		set diagram = Repository.GetCurrentDiagram
		'save the diagram first
		Repository.SaveDiagram diagram.DiagramID
		'then loop all diagramLinks
		if not diagram is nothing then
			for each diagramLink in diagram.DiagramLinks
				set connector = Repository.GetConnectorByID(diagramLink.ConnectorID)
				if not connector is nothing then
					'set the connectorstyle
					setConnectorStyle diagramLink, determineStyle(connector)
					'save the diagramlink
					diagramLink.Update
					dirty = true
				end if
			next
			'reload the diagram if we changed something
			if dirty then
				'reload the diagram to show the link style
				Repository.ReloadDiagram diagram.DiagramID
			end if
		end if
	end if
end function

'gets the diagram link object
function getdiagramLinkForConnector(connector, diagram)
	dim diagramLink 
	set getdiagramLinkForConnector = nothing
	for each diagramLink in diagram.DiagramLinks
		if diagramLink.ConnectorID = connector.ConnectorID then
			set getdiagramLinkForConnector = diagramLink
			exit for
		end if
	next
end function

'actually sets the connector style
function setConnectorStyle(diagramLink, connectorStyle)
	'split the style into its parts
	dim styleparts
	dim styleString
	styleString = diagramLink.Style
	styleparts = Split(styleString,";")
	dim stylePart
	dim mode
	dim modeIndex
	modeIndex = -1
	dim tree
	dim treeIndex
	treeIndex = -1
	mode = ""
	tree = ""
	dim i
	'find if Mode and Tree are already defined
	for i = 0 to Ubound(styleparts) -1 
		stylePart = styleparts(i)
		if Instr(stylepart,"Mode=") > 0 then
			modeIndex = i
		elseif Instr(stylepart,"TREE=") > 0 then
			treeIndex = i
		end if
	next
	'these connectorstyles use mode=3 and the tree
	if  connectorStyle = lsTreeVerticalTree or _
		connectorStyle = lsTreeHorizontalTree or _
		connectorStyle = lsLateralHorizontalTree or _
		connectorStyle = lsLateralVerticalTree or _
		connectorStyle = lsOrthogonalSquareTree or _
		connectorStyle = lsOrthogonalRoundedTree then
		mode = "3"
		tree = connectorStyle
	else
		mode = connectorStyle
	end if
	'set the mode value
	if modeIndex >= 0 then
		styleparts(modeIndex) = "Mode=" & mode
		diagramLink.Style = join(styleparts,";")
	else
		diagramLink.Style = "Mode=" & mode& ";"& diagramLink.Style
	end if
	'set the tree value
	if treeIndex >= 0 then
		if len(tree) > 0 then
			styleparts(treeIndex) = "TREE=" & tree
			diagramLink.Style = join(styleparts,";")
		else
			'remove tree part
			diagramLink.Style = replace(diagramLink.Style,styleparts(treeIndex)&";" , "")
		end if
	else
		diagramLink.Style = diagramLink.Style & "TREE=" & tree & ";"
	end if
end function

function getConnectorStyle(diagramLink)
	'split the style
	dim styleparts
	styleparts = Split(diagramLink.Style,";")
	dim stylePart
	dim mode
	dim tree
	mode = ""
	tree = ""
	for each stylepart in styleparts
		if Instr(stylepart,"Mode=") > 0 then
			mode = right(stylepart, 1)
		elseif Instr(stylepart,"TREE=") > 0 then
			tree = replace(stylepart, "TREE=", "")
		end if
	next
	if tree <> "" then
		getConnectorStyle = tree
	else
		getConnectorStyle = mode
	end if
end function