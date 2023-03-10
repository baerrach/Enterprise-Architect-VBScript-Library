'[path=\ArchiMate]
'[group=ArchiMate]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Logging.LogManager
!INC Utils.Color
!INC ArchiMate.ArchiMateElement

'https://www.archimatetool.com/resources/ "Mastering ArchiMate" scheme
Dim masteringArchiMateColourScheme
Set masteringArchiMateColourScheme = CreateObject("Scripting.Dictionary")

masteringArchiMateColourScheme.Add "Implementation Event", &HFFE0DF
masteringArchiMateColourScheme.Add "Driver", &HCDCCFF
masteringArchiMateColourScheme.Add "Application Process", &HFFFFAF
masteringArchiMateColourScheme.Add "Business Actor", &HE6FFFF
masteringArchiMateColourScheme.Add "Outcome", &HCDCCFF
masteringArchiMateColourScheme.Add "Resource", &HF4E0AA
masteringArchiMateColourScheme.Add "Representation", &HE6FFE6
masteringArchiMateColourScheme.Add "Facility", &H7DFFFF
masteringArchiMateColourScheme.Add "Gap", &HDFFFE0
masteringArchiMateColourScheme.Add "Value", &HE6FFE6
masteringArchiMateColourScheme.Add "Communication Network", &H7DFFFF
masteringArchiMateColourScheme.Add "Business Object", &HE6FFE6
masteringArchiMateColourScheme.Add "Business Function", &HFFFFE6
masteringArchiMateColourScheme.Add "Course Of Action", &HF4E0AA
masteringArchiMateColourScheme.Add "Technology Interface", &H7DFFFF
masteringArchiMateColourScheme.Add "System Software", &H7DFFFF
masteringArchiMateColourScheme.Add "Device", &H7DFFFF
masteringArchiMateColourScheme.Add "Application Service", &HFFFFAF
masteringArchiMateColourScheme.Add "Business Interaction", &HFFFFE6
masteringArchiMateColourScheme.Add "Work Package", &HFFE0DF
masteringArchiMateColourScheme.Add "Data Object", &HAFFFAF
masteringArchiMateColourScheme.Add "Stakeholder", &HCDCCFF
masteringArchiMateColourScheme.Add "Application Collaboration", &HAFFFFF
masteringArchiMateColourScheme.Add "Application Function", &HFFFFAF
masteringArchiMateColourScheme.Add "Material", &H91FF91
masteringArchiMateColourScheme.Add "Application Interaction", &HFFFFAF
masteringArchiMateColourScheme.Add "Node", &H7DFFFF
masteringArchiMateColourScheme.Add "Deliverable", &HFFE0DF
masteringArchiMateColourScheme.Add "Technology Process", &HFFFF82
masteringArchiMateColourScheme.Add "Capability", &HF4E0AA
masteringArchiMateColourScheme.Add "Goal", &HCDCCFF
masteringArchiMateColourScheme.Add "Distribution Network", &H7DFFFF
masteringArchiMateColourScheme.Add "Technology Function", &HFFFF82
masteringArchiMateColourScheme.Add "Plateau", &HDFFFE0
masteringArchiMateColourScheme.Add "Application Interface", &HAFFFFF
masteringArchiMateColourScheme.Add "Technology Event", &HFFFF82
masteringArchiMateColourScheme.Add "Contract", &HE6FFE6
masteringArchiMateColourScheme.Add "Technology Service", &HFFFF82
masteringArchiMateColourScheme.Add "Business Interface", &HE6FFFF
masteringArchiMateColourScheme.Add "Path", &H7DFFFF
masteringArchiMateColourScheme.Add "Constraint", &HCDCCFF
masteringArchiMateColourScheme.Add "Requirement", &HCDCCFF
masteringArchiMateColourScheme.Add "Application Component", &HAFFFFF
masteringArchiMateColourScheme.Add "Artifact", &H91FF91
masteringArchiMateColourScheme.Add "Business Process", &HFFFFE6
masteringArchiMateColourScheme.Add "Business Collaboration", &HE6FFFF
masteringArchiMateColourScheme.Add "Business Role", &He6ffff
masteringArchiMateColourScheme.Add "Principle", &HCDCCFF
masteringArchiMateColourScheme.Add "Application Event", &HFFFFAF
masteringArchiMateColourScheme.Add "Product", &HE6FFE6
masteringArchiMateColourScheme.Add "Business Service", &HFFFFE6
masteringArchiMateColourScheme.Add "Business Event", &HFFFFE6
masteringArchiMateColourScheme.Add "Technology Interaction", &HFFFF82
masteringArchiMateColourScheme.Add "Assessment", &HCDCCFF
masteringArchiMateColourScheme.Add "Location", &HE6FFFF
masteringArchiMateColourScheme.Add "Meaning", &HE6FFE6
masteringArchiMateColourScheme.Add "Equipment", &H7DFFFF
masteringArchiMateColourScheme.Add "Technology Collaboration", &H7DFFFF

sub applyStyleColour(myArchiMateElement)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Style Colour")

	dim stereotype, defaultColor
	dim taggedValues, tvArchimateStyleColor
	
	set taggedValues = myArchiMateElement.Element.TaggedValues
	set tvArchimateStyleColor = taggedValues.GetByName("ArchiMate::Style::Color")
	if not tvArchimateStyleColor is nothing then
		if tvArchimateStyleColor.Value = "ignore" then
			logger.INFO "Default color ignoring " & myArchiMateElement.element.name
			exit sub
		end if
	end if
	
	stereotype = myArchiMateElement.Stereotype

	if Len(stereotype) > 2 then
		stereotype = Mid(stereotype, 2, Len(stereotype)-2)
	end if

	if masteringArchiMateColourScheme.Exists(stereotype) then
		defaultColor = masteringArchiMateColourScheme(stereotype)
		if (myArchiMateElement.DiagramObject.BackgroundColor <> defaultColor) then
			myArchiMateElement.DiagramObject.BackgroundColor = defaultColor
			myArchiMateElement.DiagramObject.Update()
		end if
	else
		logger.Info "Ignoring non-ArchiMate element name=" & myArchiMateElement.element.name & " stereotype=" & myArchiMateElement.element.stereotype & " type=" & myArchiMateElement.element.type
	end if

end sub
