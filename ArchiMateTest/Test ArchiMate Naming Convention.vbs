option explicit

'[path=\ArchiMateTest]
'[group=ArchiMateTest]
 
 !INC Assert.Assert
 !INC ArchiMateTest.MockElement
 !INC ArchiMate.Naming Convention
 
sub TestApplyArchiMateNamingConventionToDataObject
	dim element
	set element = new MockElement
	
	element.Name = "default name"
	element.Stereotype = "ArchiMate_DataObject"
	
	applyArchiMateNamingConventionToElement element
	
	assertEquals "ArchiMate Naming Convention applied", "[<group>] " & vbCrLf & "default name" & " " & vbCrLf & "(Data Object)", element.Name
end sub

sub TestApplyArchiMateNamingConventionToApplicationComponent
	dim element
	set element = new MockElement
	
	element.Name = "default name"
	element.Stereotype = "ArchiMate_ApplicationComponent"
	
	applyArchiMateNamingConventionToElement element
	
	assertEquals "ArchiMate Naming Convention applied", "[<group>] " & vbCrLf & "default name" & " " & vbCrLf & "(Application Component)", element.Name
end sub

sub TestApplyArchiMateNamingConventionToElementWithConventionsAlready
	dim element
	set element = new MockElement
	
	element.Name = "[My Group] " & vbCrLf & "default name" & " " & vbCrLf & "(Application Component)"
	element.Stereotype = "ArchiMate_ApplicationComponent"
	
	applyArchiMateNamingConventionToElement element
	
	assertEquals "ArchiMate Naming Convention applied", "[My Group] " & vbCrLf & "default name" & " " & vbCrLf & "(Application Component)", element.Name
end sub

sub TestApplyArchiMateNamingConventionToElementWithConventionsAlreadyButWrongStereotype
	dim element
	set element = new MockElement
	
	element.Name = "[My Group] " & vbCrLf & "default name" & " " & vbCrLf & "(Application Component)"
	element.Stereotype = "ArchiMate_DataObject"
	
	applyArchiMateNamingConventionToElement element
	
	assertEquals "ArchiMate Naming Convention applied", "[My Group] " & vbCrLf & "default name" & " " & vbCrLf & "(Data Object)", element.Name
end sub

sub main
	TestApplyArchiMateNamingConventionToDataObject
	TestApplyArchiMateNamingConventionToApplicationComponent
	TestApplyArchiMateNamingConventionToElementWithConventionsAlready
	TestApplyArchiMateNamingConventionToElementWithConventionsAlreadyButWrongStereotype
end sub

main