'[path=\ArchiMate]
'[group=ArchiMate]

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
	
	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property

	' Stereotype property.
	Public Property Get Stereotype
	  Stereotype = m_Stereotype
	End Property

	' DiagramObject property.
	Public Property Get DiagramObject
	  set DiagramObject = m_DiagramObject
	End Property

	' Element property.
	Public Property Get Element
	  set Element = m_Element
	End Property

	Private Sub initFromElementName(elementName)
		dim rx
		set rx = new RegExp
		
		dim groupPart, namePart, typePart, optionalNewlines
		groupPart = "(\[" & "[^\]]+" & "\])?[ \r\n]*"
		namePart = "([^(\r\n]*)?"
		typePart = "[ \r\n]*(\(" & "[^)]+" & "\))?"
		optionalNewlines = "\r?\n?"
		rx.Pattern = groupPart & optionalNewlines & namePart & optionalNewlines & typePart
		rx.Multiline = True
	   		
		' Find matches.
		Dim matches
		set matches = rx.Execute(elementName)

		Dim match
		set match = matches(0)
		m_Group = match.SubMatches(0)
		m_Name = Trim(match.SubMatches(1))
		m_Stereotype = match.SubMatches(2)
	end Sub

	Public function init(diagramObject, element)
		set m_DiagramObject = diagramObject
		set m_Element = element
		initFromElementName m_Element.name		
	end function
end Class
