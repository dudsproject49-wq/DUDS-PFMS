Attribute VB_Name = "VendorRepository"
Option Explicit

Private Const VEN_MOD As String = "VendorRepository"

Public Function GetNextVendorCode() As String
    Dim lCount As Long: lCount = GetRecordCount(SHT_VENDOR)
    GetNextVendorCode = "VND" & Format$(lCount + 1, "0000")
End Function

Public Function SaveVendor(ByVal lRow As Long, ByVal sCode As String, ByVal sName As String, _
    ByVal sAddress As String, ByVal sContact As String, ByVal sPhone As String, _
    ByVal sEmail As String, ByVal sNPWP As String, ByVal sBank As String) As Boolean
    Dim v(0 To 10) As Variant
    If lRow <= 0 Then
        v(0) = GenerateGUID(): v(1) = sCode: v(2) = sName: v(3) = sAddress
        v(4) = sContact: v(5) = sPhone: v(6) = sEmail: v(7) = sNPWP
        v(8) = sBank: v(9) = gCurrentUser: v(10) = Now()
        lRow = InsertRecord(SHT_VENDOR, v)
        SaveVendor = (lRow > 0)
    Else
        v(1) = sCode: v(2) = sName: v(3) = sAddress: v(4) = sContact
        v(5) = sPhone: v(6) = sEmail: v(7) = sNPWP: v(8) = sBank
        SaveVendor = UpdateRecord(SHT_VENDOR, lRow, v)
    End If
End Function

Public Function DeleteVendor(ByVal lRow As Long) As Boolean
    DeleteVendor = DeleteRecord(SHT_VENDOR, lRow)
End Function

Public Function GetVendorList() As Variant()
    GetVendorList = GetAllRecords(SHT_VENDOR)
End Function

Public Function FindVendorByCode(ByVal sCode As String) As Long
    FindVendorByCode = FindRecord(SHT_VENDOR, COL_VEN_CODE, sCode)
End Function

Public Function GetVendorDropdown() As Collection
    Dim col As New Collection, ws As Worksheet, lRow As Long, lLast As Long
    Set ws = ThisWorkbook.Worksheets(SHT_VENDOR)
    lLast = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For lRow = 2 To lLast
        If Len(SafeConvertToString(ws.Cells(lRow, 2).Value)) > 0 Then
            col.Add SafeConvertToString(ws.Cells(lRow, 2).Value) & " - " & SafeConvertToString(ws.Cells(lRow, 3).Value)
        End If
    Next lRow
    Set GetVendorDropdown = col
End Function
