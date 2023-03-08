'[path=\ArchiMateTest]
'[group=ArchiMateTest]

Class MockElement
	Private m_Name
	Private m_Stereotype
	Private m_UpdateValue
	
	Private Sub Class_Initialize
	  m_Name = ""
	  m_Stereotype = ""
	  m_UpdateValue = True
	End Sub
	
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
	
	function Update()
		Update = m_UpdateValue
	end function
end Class