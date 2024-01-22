'[path=\ArchiMate]
'[group=ArchiMate]

'
' Convert ArchiMateStereotypes of the form "ArchiMate_<sterotype>"
' to a plain stereotype name "<stereotyp>"
function ArchiMateStereotypeToStereotype(ArchiMateStereotype)
	Dim stereotypeCamelCase, stereotypeWithSpaces
	Dim rx, match, matches, i
	Set rx = new RegExp

	rx.pattern = "ArchiMate_([\w]+)"
	rx.Global = True
	set matches = rx.Execute(ArchiMateStereotype)
	set match = matches(0)
	stereotypeCamelCase = match.SubMatches(0)
	
	' Split the stereotype up on Capital letter boundaries to form a space separated version
	rx.Pattern = "[A-Z][a-z]*"
	set matches = rx.Execute(stereotypeCamelCase)
	i = 1
	stereotypeWithSpaces = ""
	for each match in matches
		stereotypeWithSpaces = stereotypeWithSpaces & match.Value
		if i <> matches.count then
			stereotypeWithSpaces = stereotypeWithSpaces & " "
		end if
		i = i + 1
	next
	ArchiMateStereotypeToStereotype = stereotypeWithSpaces
end function

Class ArchiMateElement
	Private m_Group
	Private m_Name
	Private m_Stereotype
	Private m_DiagramObject
	Private m_Element
	
	Private Sub Class_Initialize
	  m_Group = ""
	  m_Name = ""
	  m_Stereotype = ""
	  set m_DiagramObject = nothing
	  set m_Element = nothing
	End Sub
	
	' Group property.
	Public Property Get Group
	  Group = m_Group
	End Property
	Public Property Let Group(value)
	  m_Group = value
	End Property
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property

	' Stereotype property.
	Public Property Get Stereotype
	  Stereotype = m_Stereotype
	End Property
	Public Property Let Stereotype(value)
	  m_Stereotype = value
	End Property

	' DiagramObject property.
	Public Property Get DiagramObject
	  set DiagramObject = m_DiagramObject
	End Property

	' Element property.
	Public Property Get Element
	  set Element = m_Element
	End Property
	
	' FullName property.
	Public Property Get FullName
		dim separator
		separator = vbCrLf
		if stereotype = "(Node)" then
			' Nodes names should be on one line
			separator = ""
		end if
		FullName = Group & " " & separator & Name & " " & separator & Stereotype
	End Property
	
	' Predicates
	Public Function IsArchiMate
		IsArchiMate = InStr(element.Stereotype, "ArchiMate_") <> 0
	End Function 
	
	Public Function HasGroup
		HasGroup = Group <> ""
	End Function
	
	
	Public Function StereotypePartFromElementStereotype
		dim stereotype

		stereotype = element.Stereotype
		stereotype = "(" & ArchiMateStereotypeToStereotype(stereotype) & ")"
	
		StereotypePartFromElementStereotype = stereotype
	end Function

	Private Sub initFromElementName(elementName)
		dim rx
		set rx = new RegExp
		
		dim groupPart, namePart, stereotypePart, optionalNewlines
		groupPart      = "(\[" & "[^\]]+" & "\])?[ \r\n]*"
		namePart       = "([^(\r\n]*)?"
		stereotypePart = "[ \r\n]*(\(" & "[^)]+" & "\))?"
		optionalNewlines = "\r?\n?"
		rx.Pattern = groupPart & optionalNewlines & namePart & optionalNewlines & stereotypePart
		rx.Multiline = True
	   		
		' Find matches
		Dim matches
		set matches = rx.Execute(elementName)

		if matches.Count <> 1 then
			exit sub
		end if

		Dim match
		set match = matches(0)
		m_Group = match.SubMatches(0)
		m_Name = Trim(match.SubMatches(1))
		m_Stereotype = match.SubMatches(2)
	end Sub

	Public sub init(diagramObject, element)
		set m_DiagramObject = diagramObject
		set m_Element = element
		initFromElementName m_Element.name		
	end sub
end Class
