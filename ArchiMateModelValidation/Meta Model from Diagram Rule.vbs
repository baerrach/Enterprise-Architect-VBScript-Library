'[path=\ArchiMateModelValidation]
'[group=ArchiMateModelValidation]
'
'EA-Matic
'EA-Matic: http://bellekens.com/ea-matic/

option explicit 

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
!INC ArchiMate.ArchiMateElement
!INC ArchiMateModelValidation.ArchiMateModelValidationConstants

dim logger
set logger = new LoggerClass
logger.init "ArchiMate Meta Model from Diagram Rule"

' 
' EA_OnInitializeUserRules()
' is done in <your>ModelValidationRules_LoadRules
' A new RuleID must be created before it can be used in this file.
' ModelValidation.<your>ModelValidationConstants then needs to 
' define a constant for this new rule.
'

'''''''''''''''
' Your rule should do one validation only.
' Create more rules if you need them.
' Generally your rule will only need to handle one of these events below.
' Delete all the unused event handlers.
'''''''''''''''

function EA_OnRunConnectorRule(RuleID, ConnectorID)
	if MetaModelFromDiagramRuleId <> RuleID then
		exit function
	end if

	Logger.debug "EA_OnRunConnectorRule called RuleId=" & RuleID & " ConnectorID=" & CStr(ConnectorID)
	dim project as EA.Project
	set project = Repository.GetProjectInterface()
	
	'
	' Do your rule validation here.
	' Use project.PublishResult to notify any violations.
	'
	dim connector as EA.Connector
	dim clientId, supplierId
	dim client as EA.Element
	dim supplier as EA.Element
	dim clientStereotype
	dim supplierStereotype
	dim validElements
	
	set connector = Repository.GetConnectorByID(ConnectorID)
	
	if "ArchiMate_Access" <> connector.Stereotype then
		exit function
	end if
	
	clientId = connector.ClientID
	supplierId = connector.SupplierID
	set client = Repository.GetElementByID(clientId)
	clientStereotype = ArchiMateStereotypeToStereotype(client.Stereotype)
	set supplier = Repository.GetElementByID(supplierId)
	supplierStereotype = ArchiMateStereotypeToStereotype(supplier.Stereotype)
	
	if supplierStereotype = "Artifact" then
		Logger.debug "Validating Artifact"
		set validElements = CreateObject("System.Collections.ArrayList")
		validElements.Add "Technology Service"
		if not validElements.contains(clientStereotype) then
			project.PublishResult MetaModelFromDiagramRuleId, mvError, "Artifact can only be accessed by " & Join(validElements.ToArray(), ", ")
		end if		
	elseif supplierStereotype = "Data Object" then
		Logger.debug "Validating Data Object"
		set validElements = CreateObject("System.Collections.ArrayList")
		validElements.Add "Application Service"
		validElements.Add "Application Event"
		validElements.Add "Application Process"
		validElements.Add "Application Function"
		validElements.Add "Application Interaction"
		if not validElements.contains(clientStereotype) then
			project.PublishResult MetaModelFromDiagramRuleId, mvError, "Data Object can only be accessed by " & Join(validElements.ToArray(), ", ")
		end if		
	end if
	
end function
