Attribute VB_Name = "a_Main_v1_00"
Option Explicit

'CHANGE -------@##@#@#@#______))))))
Public wb_main      As Workbook, ws_main As Worksheet
Public wb_datafile  As Workbook, ws_datafile As Worksheet
Public wb_template  As Workbook, ws_template As Worksheet
Public wb_temp      As Workbook, ws_temp As Worksheet
Public dicData      As Dictionary
Public dicDetails   As Dictionary
'


Sub SplitDataFile()
    Dim FSO                 As New FileSystemObject
    Dim appExcel            As Excel.Application
    Dim objWindow           As Window
    Dim lData_FirstRow      As Long, lData_FirstCol As Long, lData_LastRow As Long, lData_LastCol As Long
    Dim rData_FirstCell     As Range, rData_LastCell As Range, rData As Range
    Dim rData_HeaderRowCell As Range
    
    Dim lTemplate_FirstRow  As Long, lTemplate_FirstCol As Long
    Dim rTemplate_FirsCell  As Range
    
    Dim vID                 As Variant
    Dim arrTemp             As Variant
    Dim rRng                As Range
    
    Dim sOutputFolder       As String
    Dim sSavePath           As String
    Dim sPass               As String
    Dim lCounter            As Long
    
    Dim dStart              As Double, dFinish As Double, dTime As Double, dAvarag As Double
    
    Dim V
    'On Error Resume Next
   
    'Unfreezing
    Call OO(bFreez:=True)
    
    Set objWindow = ActiveWindow
    Set wb_main = ThisWorkbook
    Set ws_main = wb_main.Sheets(1)
    Set ws_template = wb_main.Sheets(2)
    Set wb_datafile = OpenWorkbook(ws_main.Range("path_datafile").Value): If wb_datafile Is Nothing Then Exit Sub
    Set ws_datafile = wb_datafile.Sheets(1)

    'Timer
    ws_main.Range("range_begin").Value = Time
    
'VARIABLES SETTING
    sOutputFolder = ws_main.Range("path_output").Value & "\": If Not FSO.FolderExists(sOutputFolder) Then Exit Sub
    
    
    lData_FirstRow = ws_main.Range("data_row").Value + 1
    lData_LastRow = ws_datafile.Columns(ws_main.Range("data_col").Value).Rows(ws_datafile.Rows.Count).End(xlUp).Row
    lData_FirstCol = ws_main.Range("data_col_start").Value
    lData_LastCol = ws_main.Range("data_col_end").Value
    Set rData_HeaderRowCell = ws_datafile.Cells(ws_main.Range("data_row").Value, ws_main.Range("data_col").Value)
    Set rData_FirstCell = ws_datafile.Cells(lData_FirstRow, 1)
    Set rData_LastCell = ws_datafile.Cells(lData_LastRow, ws_datafile.Columns(ws_datafile.UsedRange.Columns.Count).Column) 'XXXXXXXXXXXXXXXX
    Set rData = Range(rData_FirstCell, rData_LastCell)
    
    lTemplate_FirstRow = ws_main.Range("template_row").Value + 1
    lTemplate_FirstCol = ws_main.Range("template_col").Value
    
    'turn off PageBreak refresh
    ws_datafile.DisplayPageBreaks = False
    ws_main.DisplayPageBreaks = False
    ws_template.DisplayPageBreaks = False
    
'GET DATA TO DICTIONARY
    Set dicData = GetGroupData(rRange:=rData, _
                                iIdColumn:=ws_main.Range("range_split").Value, _
                                lFirstColumn:=lData_FirstCol, _
                                lLastColumn:=lData_LastCol)
                                
    Interaction.DoEvents
    Set dicDetails = CreateDetails(rRange:=ws_datafile.UsedRange, _
                                    iIdColumn:=ws_main.Range("range_split").Value, _
                                    sFileName:=ws_main.Range("range_filename").Value, _
                                    iFileNameColumn:=ws_main.Range("range_filename_addon").Value)
    'dicDetails(0)-RowNo
    'dicDetails(1)-Password
    'dicDetails(2)-FileName
    
On Error GoTo Ext
'    Set appExcel = wb_main.Parent
'    appExcel.Visible = False
    objWindow.Visible = False
    
'DATA SPLITTING
    For Each vID In dicData.Keys
        lCounter = lCounter + 1
        dStart = Timer
        
        ws_template.Copy
        Set wb_temp = ActiveWorkbook
        Set ws_temp = ActiveSheet
        
        Set rRng = ws_temp.Cells(lTemplate_FirstRow, lTemplate_FirstCol)
        
        arrTemp = TabelaDict(dicData(vID).Items)
        rRng.Resize(UBound(arrTemp, 1), UBound(arrTemp, 2)) = arrTemp
        
        Call Un_HideColumnsInTemplate(shMain:=ws_main, shTemplate:=ws_temp)
        
        sSavePath = sOutputFolder & dicDetails(vID)(2)
        If ws_main.OLEObjects("cb_password").Object.Value Then sPass = dicDetails(vID)(1) Else sPass = vbNullString
        wb_temp.SaveAs Filename:=sSavePath, _
                        FileFormat:=xlOpenXMLWorkbook, _
                        Password:=sPass
        wb_temp.Close False
        
        ws_datafile.Cells(dicDetails(vID)(0), ws_main.Range("data_password").Value) = dicDetails(vID)(1)
        ws_datafile.Cells(dicDetails(vID)(0), ws_main.Range("data_file").Value) = dicDetails(vID)(2)
        
        dFinish = Timer
        dTime = dFinish - dStart
        dAvarag = (dAvarag + dTime) / 2
'        Debug.Print "Creating template " & lCounter & " /" & dicData.Count & vbTab & "(" & Format(dTime,"0.00") & " / Avg: " & Format(dAvarag,"0.00") & ")"
        Application.StatusBar = _
            "Creating template " & lCounter & " /" & dicData.Count & " (" & Format(dTime, "0.00") & " / Avg: " & Format(dAvarag, "0.00") & ")"
        Interaction.DoEvents
    Next vID
    
    sSavePath = ThisWorkbook.Path & "\" & CLng(Timer) & "_Data_File_w_passwords.xlsx"
    wb_datafile.SaveCopyAs sSavePath

Ext:
    On Error Resume Next
    ws_main.Range("range_end").Value = Time
    
    If Err.Number = 0 Then
        MsgBox "Consolidation successfully ended." & Chr(10) & "Duration: " & ws_main.Range("range_duration").Text & " minutes"
    Else
        MsgBox "Procedure corrupted by error." & vbCrLf & Err.Number & vbCrLf & Err.Description, vbError, "Error"
    End If
    
    objWindow.Visible = True
    appExcel.Visible = True
    Set appExcel = Nothing
    
    'Unfreezing
    Call OO(bFreez:=False)
    
    On Error GoTo 0
End Sub


