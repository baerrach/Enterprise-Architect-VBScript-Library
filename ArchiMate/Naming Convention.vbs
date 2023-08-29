'[path=\ArchiMate]
'[group=ArchiMate]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Logging.LogManager

!INC Utils.Color
!INC ArchiMate.ArchiMateElement

sub applyArchiMateNamingConventionToElement(element)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Naming Convention")
	
	dim asArchiMateElement
	set asArchiMateElement = new ArchiMateElement
	asArchiMateElement.init nothing, element

	if not asArchiMateElement.IsArchiMate then
		logger.Info "Ignoring non-ArchiMate element name=" & element.name & " stereotype=" & element.stereotype & " type=" & element.type
		exit sub
	end if 

	if asArchiMateElement.Group = "" then
		asArchiMateElement.Group = "[ <group> ]"
	end if
	
	dim expectedStereotype
	expectedStereotype = asArchiMateElement.StereotypePartFromElementStereotype() 
	if asArchiMateElement.StereoType <> expectedStereotype then
		logger.INFO "Changing " & element.name & " to have stereotype=" & expectedStereotype
		asArchiMateElement.StereoType = expectedStereotype
	end if

	dim newName
	newName = asArchiMateElement.FullName

	if element.Name <> newName then
		element.Name = newName
		if not element.Update() then
			logger.ERROR "Update failed: " & element.GetLastError()
		end if
	end if

end sub
