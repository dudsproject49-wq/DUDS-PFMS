Attribute VB_Name = "BackupService"
Option Explicit

Private Const BAK_MODULE As String = "BackupService"

Public Function CreateBackup() As Boolean
    Dim sBackupPath As String, sFileName As String, sTarget As String
    On Error GoTo CreateBackup_Err
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    sBackupPath = ThisWorkbook.Path & "\" & PATH_BACKUP_FOLDER & "\" & Format$(Now, "yyyy-mm-dd")
    If Dir(sBackupPath, vbDirectory) = "" Then MkDir sBackupPath
    sFileName = APP_NAME & "_Backup_" & Format$(Now, "yyyymmdd_hhmmss") & ".xlsm"
    sTarget = sBackupPath & "\" & sFileName
    ThisWorkbook.SaveCopyAs sTarget
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
    LogInfo BAK_MODULE & ".CreateBackup", "Backup created: " & sTarget, "Backup"
    CreateBackup = True
    Exit Function
CreateBackup_Err:
    HandleError BAK_MODULE & ".CreateBackup", Err.Number, Err.Description
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
    CreateBackup = False
End Function

Public Function RestoreBackup(ByVal BackupFilePath As String) As Boolean
    Dim wb As Workbook
    On Error GoTo RestoreBackup_Err
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    If Dir(BackupFilePath) = "" Then
        MsgBox "Backup file not found: " & BackupFilePath, vbExclamation: RestoreBackup = False: GoTo RestoreBackup_Exit
    End If
    If Not ValidateBackup(BackupFilePath) Then
        MsgBox "Invalid or corrupted backup file.", vbCritical: RestoreBackup = False: GoTo RestoreBackup_Exit
    End If
    Set wb = Workbooks.Open(BackupFilePath)
    wb.SaveAs ThisWorkbook.FullName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    wb.Close
    LogInfo BAK_MODULE & ".RestoreBackup", "Restored from: " & BackupFilePath, "Backup"
    RestoreBackup = True
    GoTo RestoreBackup_Exit
RestoreBackup_Err:
    HandleError BAK_MODULE & ".RestoreBackup", Err.Number, Err.Description
    RestoreBackup = False
RestoreBackup_Exit:
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
End Function

Public Function ValidateBackup(ByVal FilePath As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(FilePath, ReadOnly:=True, Password:="")
    If Err.Number <> 0 Then ValidateBackup = False: Exit Function
    If wb.VBProject.VBComponents.Count > 0 Then
        ValidateBackup = True
    Else
        ValidateBackup = False
    End If
    wb.Close False
    On Error GoTo 0
End Function

Public Function RepairDatabase() As Boolean
    Dim ws As Worksheet, i As Long
    On Error GoTo RepairDatabase_Err
    Application.ScreenUpdating = False: Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 1) = "_" Then
            On Error Resume Next
            ws.Cells.Replace What:="#VALUE!", Replacement:="", LookAt:=xlWhole
            ws.Cells.Replace What:="#REF!", Replacement:="", LookAt:=xlWhole
            ws.Cells.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole
            ws.Cells.Replace What:="Error ", Replacement:="", LookAt:=xlPart
            On Error GoTo 0
        End If
    Next ws
    ThisWorkbook.Save
    LogInfo BAK_MODULE & ".RepairDatabase", "Database repair completed.", "Backup"
    RepairDatabase = True
    GoTo RepairDatabase_Exit
RepairDatabase_Err:
    HandleError BAK_MODULE & ".RepairDatabase", Err.Number, Err.Description
    RepairDatabase = False
RepairDatabase_Exit:
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
End Function

Public Function GetLatestBackup() As String
    Dim sBase As String, sDateFolder As String, sFile As String
    sBase = ThisWorkbook.Path & "\" & PATH_BACKUP_FOLDER & "\"
    sDateFolder = Format$(Now, "yyyy-mm-dd")
    sFile = Dir(sBase & sDateFolder & "\" & "*.xlsm")
    If Len(sFile) > 0 Then
        GetLatestBackup = sBase & sDateFolder & "\" & sFile
    Else
        GetLatestBackup = ""
    End If
End Function

Public Function GetBackupFileList() As Variant()
    Dim sBase As String, sFolder As String, sFile As String
    Dim arrResult() As String, n As Long
    sBase = ThisWorkbook.Path & "\" & PATH_BACKUP_FOLDER & "\"
    n = 0: ReDim arrResult(1 To 100)
    sFolder = Dir(sBase, vbDirectory)
    Do While Len(sFolder) > 0
        If sFolder <> "." And sFolder <> ".." Then
            If GetAttr(sBase & sFolder) And vbDirectory Then
                sFile = Dir(sBase & sFolder & "\*.xlsm")
                Do While Len(sFile) > 0
                    n = n + 1
                    If n > UBound(arrResult) Then ReDim Preserve arrResult(1 To n + 50)
                    arrResult(n) = sBase & sFolder & "\" & sFile
                    sFile = Dir()
                Loop
            End If
        End If
        sFolder = Dir()
    Loop
    If n = 0 Then GetBackupFileList = Array(): Exit Function
    ReDim Preserve arrResult(1 To n)
    GetBackupFileList = arrResult
End Function

Public Sub AutoBackup()
    If Hour(Now) >= 18 Then
        Dim sToday As String: sToday = Format$(Now, "yyyy-mm-dd")
        Dim sCheck As String: sCheck = ""
        sCheck = Dir(ThisWorkbook.Path & "\" & PATH_BACKUP_FOLDER & "\" & sToday & "\*.xlsm")
        If Len(sCheck) = 0 Then
            CreateBackup
        End If
    End If
End Sub

