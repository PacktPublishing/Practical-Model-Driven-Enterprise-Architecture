!INC Local Scripts.EAConstants-VBScript

' Hide all connectors in the current diagram that do NOT involve the selected element
' Use for hiding extra connectors on focused metamodels
sub HideUnrelatedConnectors()

    dim diagram as EA.Diagram
    set diagram = Repository.GetCurrentDiagram()

    if diagram is nothing then
        Session.Prompt "No active diagram found.", promptOK
        exit sub
    end if

    ' Ensure exactly one element is selected
    if diagram.SelectedObjects.Count <> 1 then
        Session.Prompt "Select exactly ONE element on the diagram.", promptOK
        exit sub
    end if

    dim selectedObj as EA.DiagramObject
    set selectedObj = diagram.SelectedObjects.GetAt(0)

    dim selectedElementID
    selectedElementID = selectedObj.ElementID

    dim i
    for i = 0 to diagram.DiagramLinks.Count - 1

        dim dlink as EA.DiagramLink
        set dlink = diagram.DiagramLinks.GetAt(i)

        ' Retrieve the underlying connector
        dim connector as EA.Connector
        set connector = Repository.GetConnectorByID(dlink.ConnectorID)

        ' Check if connector involves the selected element
        dim isRelated
        isRelated = (connector.ClientID = selectedElementID) OR (connector.SupplierID = selectedElementID)

        if not isRelated then
            ' Hide the connector on the diagram
            dlink.IsHidden = True
            dlink.Update()
        end if

    next

    diagram.DiagramLinks.Refresh()
    Repository.ReloadDiagram diagram.DiagramID

    Session.Prompt "Unrelated connectors hidden.", promptOK

end sub

HideUnrelatedConnectors


