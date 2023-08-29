'[path=\ArchiMateModelValidation]
'[group=ArchiMateModelValidation]
'
'EA-Matic
'EA-Matic: http://bellekens.com/ea-matic/

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Logging.Logger
!INC Logging.LogManager
' The creating of the category and rules must be done separately to
' creating the constants used by the rules themselves.
' This problem is caused by EA-Matic and scripts being reset every 5 minutes.
' See README.md for more details.
' However, including the constants here ensures creating and using are kept in sync and
' to avoid scenarios where a new rule is created but not used (and vice versa)
!INC ArchiMateModelValidation.ArchiMateModelValidationConstants

dim logger
set logger = new LoggerClass
logger.init "ArchiMateModelValidationRules_LoadRules"

' This should be
' function EA_OnInitializeUserRules()
' but that function never receives events.
' Using EA_FileOpen isntead
function EA_FileOpen()
	Logger.debug "EA_OnInitializeUserRules called"

	dim project as EA.Project

	set project = Repository.GetProjectInterface()
	archiMateCategoryId = project.DefineRuleCategory("ArchiMate Model Validation")
	Logger.debug "EA_OnInitializeUserRules archiMateCategoryId =" & archiMateCategoryId

	NamingConventionRuleId = DefineRule(project, archiMateCategoryId, mvError, "Naming Convention")
end function

function DefineRule(project, categoryId, enumMVErrorType, ruleName)
	dim ruleId
	' The second parameter uses EnumMVErrorType values which are defined in Local Scripts.EAConstants-VBScript
	' and are mvError, mvWarning, mvInformation, mvErrorCritical.
	' The third paramter is a string for the error message.
	' Both these values are provided in Project.PublishResult so why are they also needed here? 
	ruleId = project.DefineRule(categoryId, enumMVErrorType, ruleName)
	Logger.debug "EA_OnInitializeUserRules " & ruleName & "=" & ruleId

	DefineRule = ruleId
end function
