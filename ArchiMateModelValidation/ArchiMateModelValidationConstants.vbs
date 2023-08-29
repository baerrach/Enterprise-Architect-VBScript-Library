'[path=\ArchiMateModelValidation]
'[group=ArchiMateModelValidation]

' Change this to the value in System Output/Logging after ArchiMateModelValidationRules_LoadRules has run
dim BASE_ID
BASE_ID = 800000

!INC ModelValidation.Utils

dim ArchiMateCategoryId, NamingConventionRuleId

ArchiMateCategoryId         = makeId(BASE_ID, 0)
NamingConventionRuleId      = makeId(BASE_ID, 1)