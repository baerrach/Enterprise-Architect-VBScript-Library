'[path=\ArchiMate]
'[group=ArchiMate]

!INC Local Scripts.EAConstants-VBScript
!INC Logging.LogManager
!INC EA-Extensions.ConnectorEndEx

function applyStyleAccessRelationUnspecifiedNavigability(connector)
	dim logger
	set logger = LogManager.getLogger("ArchiMate.Style Access Relation Unspecified Navigability Apply")
	
	if connector.Stereotype <> "ArchiMate_Access" then
		logger.info "Connector stereotype must be 'ArchiMate_Access' and not stereotype=" & connector.Stereotype
		exit function
	end if
	
	dim connectorEnd, isDirty
	isDirty = false
	
	set connectorEnd = connector.ClientEnd
	if connectorEnd.Navigable <> ConnectorEndExtension.Navigable_Unspecified then
		connectorEnd.Navigable = ConnectorEndExtension.Navigable_Unspecified
		isDirty = true
		if not connectorEnd.Update() then
			logger.WARN "Update failed: " & connectorEnd.GetLastError()
		end if
	end if
	
	set connectorEnd = connector.SupplierEnd
	if connectorEnd.Navigable <> ConnectorEndExtension.Navigable_Unspecified then
		connectorEnd.Navigable = ConnectorEndExtension.Navigable_Unspecified
		isDirty = true
		if not connectorEnd.Update() then
			logger.WARN "Update failed: " & connectorEnd.GetLastError()
		end if
	end if

	if isDirty then
		if not connector.Update() then
			logger.WARN "Update failed: " & connector.GetLastError()
		end if
	end if
	
	applyStyleAccessRelationUnspecifiedNavigability = isDirty
end function
