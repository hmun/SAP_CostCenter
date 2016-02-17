VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCostcenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function createMultiple(pControllingArea As String, pLanguageKey As String, pTestRun As String, pData As Collection)
    On Error GoTo createMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostCenter As Object
    Dim oReturn As Object
    Dim oCostCenterList As Object
    Dim oLanguageKey As Object
    Dim aControllingArea As String
    Dim aTestRun As String

    Set boCostCenter = MySAPCon.SAPBapiControl.GetSAPObject("CostCenter")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "CreateMultiple", "Return")
    Set oCostCenterList = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "CreateMultiple", "CostCenterList")
    Set oLanguageKey = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "ChangeMultiple", "LanguageKey")
    oLanguageKey("LANGU") = aLanguageKey

    Dim aCCLRow As Object
    Dim lRow As SAPCostCenterList
    For Each lRow In pData
        Set aCCLRow = oCostCenterList.Rows.Add
        aCCLRow.Value("COSTCENTER") = lSAPFormat.unpack(lRow.COSTCENTER, 10)
        aCCLRow.Value("VALID_FROM") = lRow.VALID_FROM
        aCCLRow.Value("VALID_TO") = lRow.VALID_TO
        aCCLRow.Value("PERSON_IN_CHARGE") = lRow.PERSON_IN_CHARGE
        aCCLRow.Value("COSTCENTER_TYPE") = lRow.COSTCENTER_TYPE
        aCCLRow.Value("COSTCTR_HIER_GRP") = lRow.COSTCTR_HIER_GRP
        aCCLRow.Value("COMP_CODE") = lRow.COMP_CODE
        aCCLRow.Value("CURRENCY") = lRow.aCURRENCY
        aCCLRow.Value("PROFIT_CTR") = lRow.PROFIT_CTR
        aCCLRow.Value("NAME") = lRow.NAME
        aCCLRow.Value("DESCRIPT") = lRow.DESCRIPT
        aCCLRow.Value("FUNC_AREA") = lRow.FUNC_AREA
        aCCLRow.Value("BUS_AREA") = lRow.BUS_AREA
        aCCLRow.Value("TELCO_DATA_LINE") = lRow.TELCO_DATA_LINE
        aCCLRow.Value("LOCK_IND_ACTUAL_REVENUES") = lRow.LOCK_IND_ACTUAL_REVENUES
        aCCLRow.Value("LOCK_IND_PLAN_REVENUES") = lRow.LOCK_IND_PLAN_REVENUES
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boCostCenter.createMultiple ControllingArea:=pControllingArea, _
        TestRun:=pTestRun, _
        LanguageKey:=oLanguageKey, _
        CostCenterList:=oCostCenterList, _
        Return:=oReturn
    Else
        boCostCenter.createMultiple ControllingArea:=pControllingArea, _
        CostCenterList:=oCostCenterList, _
        LanguageKey:=oLanguageKey, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        createMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            createMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                createMultiple = createMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boCostCenter = Nothing
    Exit Function

createMultiple_Error:
    Application.Cursor = xlDefault
    createMultiple = "Error"
    MySAPErr.MSGProt "SAPCostcenter", "createMultiple", "", Err.Number, Err.Description
    Exit Function
End Function

Function changeMultiple(pControllingArea As String, pLanguageKey As String, pTestRun As String, pData As Collection)
    On Error GoTo changeMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostCenter As Object
    Dim oReturn As Object
    Dim oCostCenterList As Object
    Dim oLanguageKey As Object
    Dim aControllingArea As String
    Dim aTestRun As String

    Set boCostCenter = MySAPCon.SAPBapiControl.GetSAPObject("CostCenter")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "ChangeMultiple", "Return")
    Set oCostCenterList = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "ChangeMultiple", "CostCenterList")
    Set oLanguageKey = MySAPCon.SAPBapiControl.DimAs(boCostCenter, "ChangeMultiple", "LanguageKey")
    oLanguageKey("LANGU") = aLanguageKey

    Dim aCCLRow As Object
    Dim lRow As SAPCostCenterList
    For Each lRow In pData
        Set aCCLRow = oCostCenterList.Rows.Add
        aCCLRow.Value("COSTCENTER") = lSAPFormat.unpack(lRow.COSTCENTER, 10)
        aCCLRow.Value("VALID_FROM") = lRow.VALID_FROM
        aCCLRow.Value("VALID_TO") = lRow.VALID_TO
        aCCLRow.Value("PERSON_IN_CHARGE") = lRow.PERSON_IN_CHARGE
        aCCLRow.Value("COSTCENTER_TYPE") = lRow.COSTCENTER_TYPE
        aCCLRow.Value("COSTCTR_HIER_GRP") = lRow.COSTCTR_HIER_GRP
        aCCLRow.Value("COMP_CODE") = lRow.COMP_CODE
        aCCLRow.Value("CURRENCY") = lRow.aCURRENCY
        aCCLRow.Value("PROFIT_CTR") = lRow.PROFIT_CTR
        aCCLRow.Value("NAME") = lRow.NAME
        aCCLRow.Value("DESCRIPT") = lRow.DESCRIPT
        aCCLRow.Value("FUNC_AREA") = lRow.FUNC_AREA
        aCCLRow.Value("BUS_AREA") = lRow.BUS_AREA
        aCCLRow.Value("TELCO_DATA_LINE") = lRow.TELCO_DATA_LINE
        aCCLRow.Value("LOCK_IND_ACTUAL_REVENUES") = lRow.LOCK_IND_ACTUAL_REVENUES
        aCCLRow.Value("LOCK_IND_PLAN_REVENUES") = lRow.LOCK_IND_PLAN_REVENUES
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boCostCenter.changeMultiple ControllingArea:=pControllingArea, _
        TestRun:=pTestRun, _
        LanguageKey:=oLanguageKey, _
        CostCenterList:=oCostCenterList, _
        Return:=oReturn
    Else
        boCostCenter.changeMultiple ControllingArea:=pControllingArea, _
        CostCenterList:=oCostCenterList, _
        LanguageKey:=oLanguageKey, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        changeMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            changeMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                changeMultiple = changeMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boCostCenter = Nothing
    Exit Function

changeMultiple_Error:
    Application.Cursor = xlDefault
    changeMultiple = "Error"
    MySAPErr.MSGProt "SAPCostcenter", "changeMultiple", "", Err.Number, Err.Description
    Exit Function
End Function
