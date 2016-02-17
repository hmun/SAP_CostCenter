Attribute VB_Name = "SAPMakro"
Sub SAP_CostCenter_create()
    Dim aSAPCostcenter As New SAPCostcenter
    Dim aSAPCostCenterList As New SAPCostCenterList
    Dim aData As New Collection

    Dim aControllingArea As String
    Dim aLanguageKey As String
    Dim aTestRun As String

    Dim i As Integer
    Dim aRetStr As String

    Worksheets("Parameter").Activate
    aControllingArea = Cells(2, 2).Value
    aLanguageKey = Cells(3, 2).Value
    aTestRun = Cells(4, 2).Value

    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("Data").Activate
    i = 2
    Do
        Set aSAPCostCenterList = New SAPCostCenterList
        aSAPCostCenterList.create Cells(i, 1).Value, Cells(i, 2).Value, Cells(i, 3).Value, Cells(i, 4).Value, _
        Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value, Cells(i, 8).Value, _
        Cells(i, 9).Value, Cells(i, 10).Value, Cells(i, 11).Value, Cells(i, 12).Value, _
        Cells(i, 13).Value, Cells(i, 14), Cells(i, 15), Cells(i, 16)
        aData.Add aSAPCostCenterList
        aRetStr = aSAPCostcenter.createMultiple(aControllingArea, aLanguageKey, aTestRun, aData)
        Cells(i, 17) = aRetStr
        Set aData = New Collection
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
End Sub

Sub SAP_CostCenter_change()
    Dim aSAPCostcenter As New SAPCostcenter
    Dim aSAPCostCenterList As New SAPCostCenterList
    Dim aData As New Collection

    Dim aControllingArea As String
    Dim aLanguageKey As String
    Dim aTestRun As String

    Dim i As Integer
    Dim aRetStr As String

    Worksheets("Parameter").Activate
    aControllingArea = Cells(2, 2).Value
    aLanguageKey = Cells(3, 2).Value
    aTestRun = Cells(4, 2).Value

    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("Data").Activate
    i = 2
    Do
        Set aSAPCostCenterList = New SAPCostCenterList
        aSAPCostCenterList.create Cells(i, 1).Value, Cells(i, 2).Value, Cells(i, 3).Value, Cells(i, 4).Value, _
        Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value, Cells(i, 8).Value, _
        Cells(i, 9).Value, Cells(i, 10).Value, Cells(i, 11).Value, Cells(i, 12).Value, _
        Cells(i, 13).Value, Cells(i, 14), Cells(i, 15), Cells(i, 16)
        aData.Add aSAPCostCenterList
        aRetStr = aSAPCostcenter.changeMultiple(aControllingArea, aLanguageKey, aTestRun, aData)
        Cells(i, 17) = aRetStr
        Set aData = New Collection
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
End Sub
