option explicit

'[path=\ArchiMate]
'[group=ArchiMate]
'
'EA-Matic
'EA-Matic: http://bellekens.com/ea-matic/

!INC Local Scripts.EAConstants-VBScript
!INC Utils.Util
!INC Logging.LogManager
!INC ArchiMate.Style Access Relation Unspecified Navigability Apply

' https://sparxsystems.com/enterprise_architect_user_guide/15.2/automation/broadcastpostnewconnector.html
' Info.ConnectorID: A long value corresponding to Connector.ConnectorID
' Return True if the connector has been updated during this notification. Return False otherwise.
function EA_OnPostNewConnector(Info)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.EA-Matic Apply ArchiMate Conventions On New Diagram Object")
	
	logger.INFO "Start..."

	'''gather required details
	dim connectorId, diagram
	connectorId = Info.Get("ConnectorID")
	logger.DEBUG " ConnectorID=" & ConnectorID

	dim connector as EA.Connector
	set connector = Repository.GetConnectorByID (connectorId)

	set diagram = Repository.GetCurrentDiagram()

	'''validate details and provide error message

	'''Delegate actual work to included script
	logger.info "Working on connector.ConnectorID=" & connector.ConnectorID
	EA_OnPostNewConnector = applyStyleAccessRelationUnspecifiedNavigability(connector)

	logger.INFO "Done"
end function