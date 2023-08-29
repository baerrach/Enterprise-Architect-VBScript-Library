'[path=\ArchiMateModelValidation]
'[group=ArchiMateModelValidation]
'
'EA-Matic
'EA-Matic: http://bellekens.com/ea-matic/

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
!INC ArchiMateModelValidation.ArchiMateModelValidationConstants
!INC ArchiMate.ArchiMateElement

dim logger
set logger = new LoggerClass
logger.init "ArchiMateModelValidationRule_<Name>"

'
' EA_OnInitializeUserRules()
' is done in <your>ModelValidationRules_LoadRules
' A new RuleID must be created before it can be used in this file.
' ModelValidation.ArchiMateModelValidationConstants then needs to
' define a constant for this new rule.
'

'''''''''''''''
' Your rule should do one validation only.
' Create more rules if you need them.
' Generally your rule will only need to handle one of these events below.
' Delete all the unused event handlers.
'''''''''''''''

function EA_OnRunElementRule(RuleID, Element)
	Logger.debug "EA_OnRunElementRule called NamingConventionRuleId=" & RuleID & " Element.Name=" & Element.Name

	if NamingConventionRuleId <> RuleID then
		exit function
	end if

	Logger.debug "EA_OnRunElementRule called NamingConventionRuleId=" & RuleID & " Element.Name=" & Element.Name
	dim project as EA.Project
	set project = Repository.GetProjectInterface()

	'
	' Do your rule validation here.
	' Use project.PublishResult to notify any violations.
	'
	dim asArchiMateElement
	set asArchiMateElement = new ArchiMateElement
	asArchiMateElement.init nothing, Element

	if not asArchiMateElement.IsArchiMate() then
		exit function
	end if

	if not asArchiMateElement.HasGroup() then
		project.PublishResult NamingConventionRuleId, mvError, "Name is missing Group e.g. []: " & Element.Name	
	end if

	if asArchiMateElement.Stereotype <> asArchiMateElement.StereotypePartFromElementStereotype() then
		project.PublishResult NamingConventionRuleId, mvError, "Stereotype in name " & asArchiMateElement.Stereotype & " does not match ArchiMate Element stereotype " & asArchiMateElement.StereotypePartFromElementStereotype()
	end if
end function
