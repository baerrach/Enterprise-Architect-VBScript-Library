'[path=Dependency Analysis]
'[group=Dependency Analysis]

' Based on https://community.sparxsystems.com/community-resources/493-65visualization-of-package-dependencies

Option Explicit

'GLOBAL VARIABLES
Dim AddedCount, RemovedCount, SyncedCount

Sub Main()
	Session.Output "Synchronizing Package Dependency Diagram: """ & GetCurrentDiagram.Name & """"

	AddedCount = 0
	RemovedCount = 0
	SyncedCount = 0

	Dim DependenciesArray()
	ReDim DependenciesArray(0)
	
	Session.Output "Scanning for Dependencies..."
	GetDependendantPackages GetTreeSelectedPackage, DependenciesArray
	Session.Output "Synchronizing Diagram..."
	SynchronizeDiagram GetCurrentDiagram, DependenciesArray

	Session.Output SyncedCount & " Packages Synchronized, " _
		& AddedCount & " Packages Added, " _
		& RemovedCount & " Packages Removed."

	'Refresh Diagram if changes have occurred.
	If AddedCount + RemovedCount > 0 Then
		ReloadDiagram GetCurrentDiagram.DiagramID
	End If

	Session.Output "DONE."

End Sub

'Return Array of IDs for Packages with interdependencies
'Array should be initialized to zero before first call.  i.e.
' Dim DependenciesArray()
' ReDim DependenciesArray(0)
'Array is passed By Reference, so no need for a return value
Sub GetDependendantPackages(Package, DependenciesArray)
	Dim Connector, SubPackage
	
	'Session.Output "... " & Package.Name
	
	For Each Connector In Package.Connectors
		If Connector.Type = "Dependency" Then
			'Add Package ID to array
			AddToArrayIfNotExist Package.Element.ElementID, DependenciesArray
			Exit For
		End If
	Next
	Set Connector = Nothing

	SyncedCount = SyncedCount + 1

	For Each SubPackage in Package.Packages
		GetDependendantPackages SubPackage, DependenciesArray
	Next
	Set SubPackage = Nothing
	
	'GetDependendantPackages = PackageDependencies
End Sub

Sub SynchronizeDiagram(Diagram, DependenciesArray)
	Dim DiagramObject, Element, i
	Dim ExistingElements()
	
	ReDim ExistingElements(0)

	For i = 0 To Diagram.DiagramObjects.Count - 1
		Set DiagramObject = Diagram.DiagramObjects.GetAt(i)
		Set Element = GetElementByID(DiagramObject.ElementID)
		If Element.Type = "Package" Then
			If ArrayContains(DependenciesArray, DiagramObject.ElementID) Then
				AddToArrayIfNotExist DiagramObject.ElementID, ExistingElements
			Else
				Diagram.DiagramObjects.Delete i
				RemovedCount = RemovedCount + 1
			End If
		End If
		Set DiagramObject = Nothing
		Set Element = Nothing
	Next
	
	For i = 1 To UBound(DependenciesArray)
		If Not ArrayContains(ExistingElements, DependenciesArray(i)) Then
			Set DiagramObject = Diagram.DiagramObjects.AddNew("", "")
			DiagramObject.ElementID = DependenciesArray(i)
			DiagramObject.Update
			AddedCount = AddedCount + 1
		End If
	Next
	
	Diagram.DiagramObjects.Refresh

End Sub


'Return True if item was added, False if an identical value already exists
Function AddToArrayIfNotExist(Item, TheArray)
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
