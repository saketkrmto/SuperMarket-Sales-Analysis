Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
strSource = fso.GetAbsolutePathName("SUPER MARKET DATA.xlsx")
strTarget = fso.GetAbsolutePathName("SuperMarket_Dashboard.xlsm")

On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(strSource)
If Err.Number <> 0 Then
    WScript.Echo "Error opening source workbook: " & Err.Description
    objExcel.Quit
    WScript.Quit
End If
On Error GoTo 0

' Try to find Dashboard sheet, delete it if it exists
On Error Resume Next
Set wsExists = objWorkbook.Sheets("Dashboard")
If Not wsExists Is Nothing Then
    wsExists.Delete
End If
On Error GoTo 0

' Add a new worksheet for Dashboard
Set objWorksheet = objWorkbook.Sheets.Add()
objWorksheet.Name = "Dashboard"
objWorksheet.Range("A1").Value = "Supermarket Sales Dashboard"
objWorksheet.Range("A1").Font.Bold = True
objWorksheet.Range("A1").Font.Size = 16

' Add VBA Macro
On Error Resume Next
Set xlmodule = objWorkbook.VBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
If Err.Number <> 0 Then
    WScript.Echo "Warning: Could not add VBA module programmatically: " & Err.Description
Else
    strCode = "Sub RefreshDashboard()" & vbCrLf & _
              "    Dim wsData As Worksheet, wsDash As Worksheet" & vbCrLf & _
              "    Set wsDash = ThisWorkbook.Sheets(""Dashboard"")" & vbCrLf & _
              "    On Error Resume Next" & vbCrLf & _
              "    Set wsData = ThisWorkbook.Sheets(2)" & vbCrLf & _
              "    On Error GoTo 0" & vbCrLf & _
              "    If Not wsData Is Nothing Then" & vbCrLf & _
              "        wsDash.Range(""A3"").Value = ""Total Rows in Data: "" & wsData.Cells(wsData.Rows.Count, 1).End(-4162).Row" & vbCrLf & _
              "    End If" & vbCrLf & _
              "    wsDash.Range(""A4"").Value = ""Last Refreshed: "" & Now" & vbCrLf & _
              "    MsgBox ""Dashboard data refreshed successfully!"", vbInformation, ""Success""" & vbCrLf & _
              "End Sub"
    xlmodule.CodeModule.AddFromString strCode
    
    Set btn = objWorksheet.Buttons.Add(10, 50, 120, 30)
    btn.Caption = "Refresh Dashboard"
    btn.OnAction = "RefreshDashboard"
End If
On Error GoTo 0

objWorkbook.SaveAs strTarget, 52
objWorkbook.Close False
objExcel.Quit

WScript.Echo "Dashboard generation completed successfully!"
