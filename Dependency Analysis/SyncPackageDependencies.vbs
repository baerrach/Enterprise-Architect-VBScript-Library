'[path=Dependency Analysis]
'[group=Dependency Analysis]

' Based on https://community.sparxsystems.com/community-resources/493-65visualization-of-package-dependencies

Option Explicit

!INC Logging.LogManager

Dim logger

'GLOBAL VARIABLES
Dim CountOfDependenciesAdded, CountOfDependenciesRemoved, CountOfPackagesSynchronized
Dim ConnectorTypesToAnalyze

' ./Testing/DotNET Test.vbs:      set dictionary = CreateObject("Scripting.Dictionary")
Sub Main()
	set logger = LogManager.getLogger("SyncPackageDependencies")

	logger.info "Synchronizing Package Dependencies for """ & GetTreeSelectedPackage.Name & """"
	
	CountOfDependenciesAdded = 0
	CountOfDependenciesRemoved = 0
	CountOfPackagesSynchronized = 0

	set ConnectorTypesToAnalyze = CreateObject("System.Collections.ArrayList")
	ConnectorTypesToAnalyze.Add "Dependency"
	ConnectorTypesToAnalyze.Add "Generalization"


	SyncPackageDependencies "", GetTreeSelectedPackage

	logger.info _
	      CountOfPackagesSynchronized & " Packages Synchronized, " _
		& CountOfDependenciesAdded & " Dependencies Added, " _
		& CountOfDependenciesRemoved & " Dependencies Removed."
		
	logger.info "DONE."
End Sub

' Recursively synchronize dependency connectors on this package based on it's contained elements
'
' Path as String = the path from the root of the selection (empty string for root)
' Package as Package = the package to synchronze package dependencies
Sub SyncPackageDependencies(Path, Package)
	Dim fullyQualifiedName
	
	if Path = "" then
		fullyQualifiedName = Package.Name
	else 
		fullyQualifiedName = Path & "." & Package.Name
	end if
	logger.info "... " & fullyQualifiedName 
	
	Dim PackageDependencies
	Dim ActualDependencies
	Dim PackageDependencyID
	Dim i, indexIntoConnectors
	Dim Element, Connector, TargetElement, NewDependency
	Dim RefreshRequired
	
	'dependency connectors currently drawn between the packages
	PackageDependencies = GetPackageDependencies(Package)
	
	'actual dependencies that should exist based on package contents
	ActualDependencies = GetActualPackageDependencies(Package)
	
	'remove any old dependencies that no longer apply
	For i = 1 To UBound(PackageDependencies)
		logger.debug "    Checking " & Package.Name & " dependencies"

		PackageDependencyID = PackageDependencies(i)
		'does this dependency connector still correspond to an actual dependency?
		If Not ArrayContains(ActualDependencies, PackageDependencyID) Then
			Set TargetElement = GetElementByID(PackageDependencyID)
			
			'find and delete the dependency
			For indexIntoConnectors = 0 To Package.Connectors.Count - 1
				Set Connector = Package.Connectors.GetAt(indexIntoConnectors)
				If Connector.Type = "Dependency" Then
					logger.debug "    ... to  " & TargetElement.Name _
						& " Connector.ClientID=" & Connector.ClientID _
						& " Package.Element.ElementID=" & Package.Element.ElementID _
						& " Connector.SupplierID=" & PackageDependencyID
					If Connector.ClientID = Package.Element.ElementID And Connector.SupplierID = PackageDependencyID Then
						logger.info "    Removing Dependency from " & Package.Name & " to " & TargetElement.Name
						Package.Connectors.Delete indexIntoConnectors
						CountOfDependenciesRemoved = CountOfDependenciesRemoved + 1
						RefreshRequired = True
					End If
				End If
			Next

			'refresh collection after any deletions
			If RefreshRequired Then
				Package.Connectors.Refresh
				RefreshRequired = False
			End If
			
			Set Connector = Nothing
			Set TargetElement = Nothing
		End If
	Next

	'add dependencies
	For i = 1 To UBound(ActualDependencies)
		'is this actual dependency currently represented on this package?
		PackageDependencyID = ActualDependencies(i)
		If Not ArrayContains(PackageDependencies, PackageDependencyID) Then
			Set TargetElement = GetElementByID(PackageDependencyID)
			CountOfDependenciesAdded = CountOfDependenciesAdded + 1
			logger.info "    Adding Dependency from " & Package.Name & " to " & TargetElement.Name
			
			Set NewDependency = Package.Connectors.AddNew("", "Dependency")
			NewDependency.SupplierID = PackageDependencyID
			NewDependency.Update

			Set TargetElement = Nothing
			Set NewDependency = Nothing
		End If
	Next
	
	CountOfPackagesSynchronized = CountOfPackagesSynchronized + 1
	
	'Synchronize the subpackages
	Dim SubPackage
	For Each SubPackage In Package.Packages
		SyncPackageDependencies fullyQualifiedName, SubPackage
	Next
	
End Sub


'Return Array of ID's for direct dependencies currently from this package
Function GetPackageDependencies(Package)
	Dim PackageDependencies()
	Dim Connector, TargetElement
	
	ReDim PackageDependencies(0)
	For Each Connector In Package.Connectors
		If Connector.Type = "Dependency" Then
			If Connector.ClientID = Package.Element.ElementID Then
				Set TargetElement = GetElementByID(Connector.SupplierID)
				If TargetElement.Type = "Package" Then
					Logger.debug "GetPackageDependencies: Found " & TargetElement.Name 
					AddToArrayIfNotExist PackageDependencies, Connector.SupplierID
				End If
				Set TargetElement = Nothing
			End If
		End If
	Next
	
	Set Connector = Nothing
	
	GetPackageDependencies = PackageDependencies
End Function

'Return Array of ID's for dependencies of this package based on it's contained elements
Function GetActualPackageDependencies(Package)
	Dim PackageDependencies()
	
	Dim Element, Connector, TargetElement, TargetElementsPackage
	
	ReDim PackageDependencies(0)

	For Each Element In Package.Elements
		If Element.Type = "Class" Or Element.Type = "Interface" Then
			logger.debug "    - Checking Element " & Element.Name
			For Each Connector In Element.Connectors
				logger.debug "        - Checking connector " & Connector.Type
			
				If ConnectorTypesToAnalyze.Contains(Connector.Type) Then
					'collect foreign dependencies only
					If Connector.ClientID = Element.ElementID Then
						Set TargetElement = GetElementByID(Connector.SupplierID)
						If TargetElement.PackageID <> Element.PackageID Then
							' Dependencies are to the Package as an Element
							Set TargetElementsPackage = GetPackageByID(TargetElement.PackageID)
							
							Logger.debug "GetActualPackageDependencies: Found " & TargetElementsPackage.Name
							
							AddToArrayIfNotExist PackageDependencies, TargetElementsPackage.Element.ElementID
						End If
						Set TargetElement = Nothing
					End If
				End If
			Next
		End If
	Next

	Set Element = Nothing
	Set Connector = Nothing
	
	GetActualPackageDependencies = PackageDependencies
End Function

' Replace VBArrays with
' CreateObject("System.Collections.ArrayList")
' Then delete these functions

'Return True if item was added, False if an identical value already exists
Function AddToArrayIfNotExist(TheArray, Item)
	Dim i
	For i = 1 To UBound(TheArray)
		If Item = TheArray(i) Then
			AddToArrayIfNotExist = False
			Exit Function
		End If
	Next
	ReDim Preserve TheArray(UBound(TheArray) + 1)
	TheArray(UBound(TheArray)) = Item
	AddToArrayIfNotExist = True
End Function

Function ArrayContains(TheArray, Item)
	Dim i
	For i = 1 To UBound(TheArray)
		If Item = TheArray(i) Then
			ArrayContains = True
			Exit Function
		End If
	Next
	ArrayContains = False
End Function

Main
