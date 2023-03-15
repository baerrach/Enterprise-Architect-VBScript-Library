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
	
	dim stereotype
	stereotype = element.Stereotype
	
	if InStr(1, stereotype, "ArchiMate_") = 0 then
		logger.Info "Ignoring non-ArchiMate element name=" & element.name & " stereotype=" & element.stereotype & " type=" & element.type
		exit sub
	end if 
	
	dim asArchiMateElement
	set asArchiMateElement = new ArchiMateElement
	asArchiMateElement.init nothing, element
	
	Dim rx, match, matches, i
	Set rx = CreateObject("VBScript.RegExp")

	rx.pattern = "ArchiMate_([\w]+)"
    rx.Global = True
	set matches = rx.Execute(stereotype)
	set match = matches(0)
	stereotype = match.SubMatches(0)
	
	' Split the stereotype up on Capital letter boundaries to form a space separated version
	rx.Pattern = "[A-Z][a-z]*"
	set matches = rx.Execute(stereotype)
	stereotype = "("
	i = 1
	for each match in matches
		stereotype = stereotype & match.Value
		if i <> matches.count then
			stereotype = stereotype & " "
		end if
		i = i + 1
	next
	stereotype = stereotype & ")"
	
	if asArchiMateElement.Group = "" then
		asArchiMateElement.Group = "[ <group> ]"
	end if
	
	if asArchiMateElement.StereoType <> "" and asArchiMateElement.StereoType <> stereotype then
		logger.INFO "Changing " & element.name & " to have stereotype=" & stereotype
	end if
	asArchiMateElement.StereoType = stereotype

	dim newName
	newName = asArchiMateElement.Group & " " & vbCrLf & asArchiMateElement.Name & " " & vbCrLf & asArchiMateElement.stereotype

	if element.Name <> newName then
		element.Name = newName
		if not element.Update() then
			logger.ERROR element.GetLastError()
		end if
	end if

end sub
