Attribute VB_Name = "ProjectRepository"
Option Explicit

Private Const TBL_NAME As String = "tblProject"

Public Function CreateProject(ByVal ProjectID As String, _
                              ByVal ProjectCode As String, _
                              ByVal ProjectName As String, _
                              ByVal ClientName As String, _
                              ByVal Location As String, _
                              ByVal StartDate As Date, _
                              ByVal EndDate As Date, _
                              ByVal ContractValue As Double, _
                              ByVal BudgetValue As Double, _
                              ByVal Status As String, _
                              ByVal Progress As Double, _
                              ByVal Notes As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lRow As Long

    On Error GoTo CreateProject_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    lRow = lo.ListRows.AddAlwaysNewRow.Position

    With lo.DataBodyRange
        .Cells(lRow, COL_PROJ_ID + 1).Value = ProjectID
        .Cells(lRow, COL_PROJ_CODE + 1).Value = ProjectCode
        .Cells(lRow, COL_PROJ_NAME + 1).Value = ProjectName
        .Cells(lRow, COL_PROJ_CLIENT + 1).Value = ClientName
        .Cells(lRow, COL_PROJ_LOCATION + 1).Value = Location
        .Cells(lRow, COL_PROJ_START_DATE + 1).Value = StartDate
        .Cells(lRow, COL_PROJ_END_DATE + 1).Value = EndDate
        .Cells(lRow, COL_PROJ_CONTRACT_VAL + 1).Value = ContractValue
        .Cells(lRow, COL_PROJ_BUDGET + 1).Value = BudgetValue
        .Cells(lRow, COL_PROJ_STATUS + 1).Value = Status
        .Cells(lRow, COL_PROJ_PROGRESS + 1).Value = Progress
        .Cells(lRow, COL_PROJ_NOTES + 1).Value = Notes
        .Cells(lRow, COL_PROJ_CREATED_BY + 1).Value = gCurrentUser
        .Cells(lRow, COL_PROJ_CREATED_ON + 1).Value = Now()
    End With

    CreateProject = True
    Exit Function

CreateProject_Err:
    LogError "ProjectRepository.CreateProject", Err.Description
    CreateProject = False
End Function

Public Function UpdateProject(ByVal ProjectID As String, _
                              ByVal ProjectCode As String, _
                              ByVal ProjectName As String, _
                              ByVal ClientName As String, _
                              ByVal Location As String, _
                              ByVal StartDate As Date, _
                              ByVal EndDate As Date, _
                              ByVal ContractValue As Double, _
                              ByVal BudgetValue As Double, _
                              ByVal Status As String, _
                              ByVal Progress As Double, _
                              ByVal Notes As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rFound As Range
    Dim lDataRow As Long

    On Error GoTo UpdateProject_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    Set rFound = lo.DataBodyRange.Columns(COL_PROJ_ID + 1).Find(What:=ProjectID, LookAt:=xlWhole)
    If rFound Is Nothing Then
        UpdateProject = False
        Exit Function
    End If

    lDataRow = rFound.Row - lo.HeaderRowRange.Row

    With lo.DataBodyRange
        .Cells(lDataRow, COL_PROJ_CODE + 1).Value = ProjectCode
        .Cells(lDataRow, COL_PROJ_NAME + 1).Value = ProjectName
        .Cells(lDataRow, COL_PROJ_CLIENT + 1).Value = ClientName
        .Cells(lDataRow, COL_PROJ_LOCATION + 1).Value = Location
        .Cells(lDataRow, COL_PROJ_START_DATE + 1).Value = StartDate
        .Cells(lDataRow, COL_PROJ_END_DATE + 1).Value = EndDate
        .Cells(lDataRow, COL_PROJ_CONTRACT_VAL + 1).Value = ContractValue
        .Cells(lDataRow, COL_PROJ_BUDGET + 1).Value = BudgetValue
        .Cells(lDataRow, COL_PROJ_STATUS + 1).Value = Status
        .Cells(lDataRow, COL_PROJ_PROGRESS + 1).Value = Progress
        .Cells(lDataRow, COL_PROJ_NOTES + 1).Value = Notes
    End With

    UpdateProject = True
    Exit Function

UpdateProject_Err:
    LogError "ProjectRepository.UpdateProject", Err.Description
    UpdateProject = False
End Function

Public Function DeleteProject(ByVal ProjectID As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rFound As Range

    On Error GoTo DeleteProject_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    Set rFound = lo.DataBodyRange.Columns(COL_PROJ_ID + 1).Find(What:=ProjectID, LookAt:=xlWhole)
    If rFound Is Nothing Then
        DeleteProject = False
        Exit Function
    End If

    rFound.EntireRow.Delete
    DeleteProject = True
    Exit Function

DeleteProject_Err:
    LogError "ProjectRepository.DeleteProject", Err.Description
    DeleteProject = False
End Function

Public Function FindProject(ByVal SearchColumn As Long, _
                            ByVal SearchValue As Variant) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rFound As Range

    On Error GoTo FindProject_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    If lo.DataBodyRange Is Nothing Then
        FindProject = Empty
        Exit Function
    End If

    Set rFound = lo.DataBodyRange.Columns(SearchColumn + 1).Find(What:=SearchValue, LookAt:=xlWhole)
    If rFound Is Nothing Then
        FindProject = Empty
        Exit Function
    End If

    FindProject = Application.Index(lo.DataBodyRange, rFound.Row - lo.HeaderRowRange.Row, 0)
    Exit Function

FindProject_Err:
    LogError "ProjectRepository.FindProject", Err.Description
    FindProject = Empty
End Function

Public Function GetProject(ByVal RowNum As Long) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo GetProject_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    If lo.DataBodyRange Is Nothing Then
        GetProject = Array()
        Exit Function
    End If

    GetProject = Application.Index(lo.DataBodyRange, RowNum, 0)
    Exit Function

GetProject_Err:
    LogError "ProjectRepository.GetProject", Err.Description
    GetProject = Array()
End Function

Public Function LoadProjects() As Variant
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error GoTo LoadProjects_Err

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    If lo.DataBodyRange Is Nothing Then
        LoadProjects = Array()
        Exit Function
    End If

    LoadProjects = lo.DataBodyRange.Value
    Exit Function

LoadProjects_Err:
    LogError "ProjectRepository.LoadProjects", Err.Description
    LoadProjects = Array()
End Function

Public Function GenerateProjectCode() As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lLastRow As Long
    Dim lNextNum As Long

    On Error GoTo GenerateProjectCode_Err

    lNextNum = 1

    Set ws = ThisWorkbook.Worksheets(SHT_PROJECTS)
    Set lo = ws.ListObjects(TBL_NAME)

    If Not lo.DataBodyRange Is Nothing Then
        lLastRow = lo.DataBodyRange.Rows.Count
        If lLastRow > 0 Then
            lNextNum = lLastRow + 1
        End If
    End If

    GenerateProjectCode = "PRJ-" & Format$(lNextNum, "0000")
    Exit Function

GenerateProjectCode_Err:
    LogError "ProjectRepository.GenerateProjectCode", Err.Description
    GenerateProjectCode = "PRJ-0001"
End Function

