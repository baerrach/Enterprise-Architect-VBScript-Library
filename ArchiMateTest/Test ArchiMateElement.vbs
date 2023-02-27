'[path=\ArchiMate]
'[group=ArchiMateTest]

!INC ArchiMateElement
!INC ArchiMate.Style Colour Apply

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

Class MockElement
	Private m_Name

	Private Sub Class_Initialize
	  m_Name = ""
	End Sub

	' Name property.
	Public Property Get Name
	  Name = m_Name
	End Property
	Public Property Let Name(value)
	  m_Name = value
	End Property
end Class

sub isEqual(expectedGroup, expectedName, expectedStereotype, actualArchiMateElement)
	if expectedGroup <> actualArchiMateElement.Group then
		Err.Raise vbObjectError + 1, "isEqual.group", "expected Group=" & expectedGroup & " but was actual=" & actualArchiMateElement.Group
	end if
	if expectedName <> actualArchiMateElement.Name then
		Err.Raise vbObjectError + 1, "isEqual.name", "expected Name=" & expectedName & " but was actual=" & actualArchiMateElement.Name
	end if
	if expectedStereotype <> actualArchiMateElement.Stereotype then
		Err.Raise vbObjectError + 1, "isEqual.stereotype", "expected Stereotype=" & expectedStereotype & " but was actual=" & actualArchiMateElement.Stereotype
	end if
end sub

function createArchiMateElement(name)
	dim OUT, diagramObject, element

	set element = new MockElement
	element.name = name
	set diagramObject = new MockDiagramObject
	set OUT = new ArchiMateElement
	OUT.init diagramObject, element

	set createArchiMateElement = OUT
end function

sub TestGroupOnly
	isEqual "[winsrv2/rdbms1]", "", "", createArchiMateElement("[winsrv2/rdbms1]")
end Sub

sub TestNameOnly
	isEqual "", "SQL Server 2016 DB", "", createArchiMateElement("SQL Server 2016 DB")
end Sub

sub TestStereotypeOnly
	isEqual "", "", "(Artifact)", createArchiMateElement("(Artifact)")
end Sub

sub TestNameContainsGroupNameStereoType
	isEqual "[winsrv2/rdbms1]", "SQL Server 2016 DB", "(Artifact)", createArchiMateElement("[winsrv2/rdbms1] SQL Server 2016 DB (Artifact)")
	isEqual "[winsrv2]", "SQL*Server System", "(Node)", createArchiMateElement("[winsrv2] SQL*Server System (Node)")
	isEqual "[winsrv2]", "SQL Server 2016", "(Technology Service)", createArchiMateElement("[winsrv2] SQL Server 2016 (Technology Service)")
	isEqual "[winsrv2]", "Windows Server 2012", "(System Software)", createArchiMateElement("[winsrv2] Windows Server 2012 (System Software)")
	isEqual "[winsrv2]", "SQL Server 2016", "(System Software)", createArchiMateElement("[winsrv2] SQL Server 2016 (System Software)")
	isEqual "[winsrv2]", "x86 Computer", "(Device)", createArchiMateElement("[winsrv2] x86 Computer (Device)")
end Sub

sub TestNameContainsGroupNameStereoTypeAndNewLines
	isEqual "[winsrv2/rdbms1]", "SQL Server 2016 DB", "(Artifact)", createArchiMateElement("[winsrv2/rdbms1] " & vbCrLf & "SQL Server 2016 DB " & vbCrLf & "(Artifact)")
	isEqual "[winsrv2]", "SQL*Server System", "(Node)", createArchiMateElement("[winsrv2] " & vbCrLf & "SQL*Server System " & vbCrLf & "(Node)")
	isEqual "[winsrv2]", "SQL Server 2016", "(Technology Service)", createArchiMateElement("[winsrv2] " & vbCrLf & "SQL Server 2016 " & vbCrLf & "(Technology Service)")
	isEqual "[winsrv2]", "Windows Server 2012", "(System Software)", createArchiMateElement("[winsrv2] " & vbCrLf & "Windows Server 2012 " & vbCrLf & "(System Software)")
	isEqual "[winsrv2]", "SQL Server 2016", "(System Software)", createArchiMateElement("[winsrv2] " & vbCrLf & "SQL Server 2016 " & vbCrLf & "(System Software)")
	isEqual "[winsrv2]", "x86 Computer", "(Device)", createArchiMateElement("[winsrv2] " & vbCrLf & "x86 Computer " & vbCrLf & "(Device)")
end Sub


sub main
	TestGroupOnly
	TestNameOnly
	TestStereotypeOnly
	TestNameContainsGroupNameStereotype
	TestNameContainsGroupNameStereoTypeAndNewLines
end sub