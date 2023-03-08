'[path=\ArchiMateTest]
'[group=ArchiMateTest]
 
Class MockDiagramObject
	private m_Element
	
	Private Sub Class_Initialize
	  set m_Element = nothing
	End Sub	
	
	' Element property.
	Public Property Get Element
	  set Element = m_Element
	End Property
	Public Property Set Element(value)
	  set m_Element = value
	End Property
end class