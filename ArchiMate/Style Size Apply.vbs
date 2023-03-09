'[path=\ArchiMate]
'[group=ArchiMate]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Logging.LogManager
!INC ArchiMate.ArchiMateElement

sub applyStyleSize(myArchiMateElement
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Style Size")

	dim defaultWidth, defaultHeight
	defaultWidth  = 150
	defaultHeight =  70

	dim taggedValues, tvArchimateStyleSize
	
	set taggedValues = myArchiMateElement.Element.TaggedValues
	set tvArchimateStyleSize = taggedValues.GetByName("ArchiMate::Style::Size")
	if not tvArchimateStyleSize is nothing then
		if tvArchimateStyleSize.Value = "ignore" then
			logger.INFO "Default size ignoring " & myArchiMateElement.element.name
			exit sub
		end if
	end if
	
	' https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/diagramobjects.html
	' there is no size or width and height, it must be calculated
	' additionally see Top: 
	'    Enterprise Architect uses a cartesian coordinate system, with {0,0} being the top-left corner of the diagram.
	'	 For this reason, Y-axis values (Top and Bottom) should always be negative.
	dim actualWidth, actualHeight
	actualWidth  = myArchiMateElement.DiagramObject.Left - myArchiMateElement.DiagramObject.Right
	actualHeight = myArchiMateElement.DiagramObject.Top - myArchiMateElement.DiagramObject.Bottom
	
	if actualWidth <> defaultWidth or actualHeight <> defaultHeight then
		myArchiMateElement.DiagramObject.Right = myArchiMateElement.DiagramObject.Left + defaultWidth
		myArchiMateElement.DiagramObject.Bottom = myArchiMateElement.DiagramObject.Top - defaultHeight
		myArchiMateElement.DiagramObject.Update()
	end if

end sub
