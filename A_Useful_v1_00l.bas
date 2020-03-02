Attribute VB_Name = "A_Global"
'<?>======================================================================================================================================================
'Create by:     Biel, Dawid[T596493]
'Updated on:    2018-04-30 11:39
'
'Contents of 'A_Global':
'1  Sub         ClearTable
'2  Function    OpenWorkbook
'3  Function    FindTextInRange
'4  Function    GetListObjectLastRowCell
'5  Sub         PasteArray
'6  Sub         test_ExcludeUnsavedCharacters
'7  Function    ExcludeUnsavedCharacters
'8  Function    DictToArr
'9  Function    ExtendArrayTable
'
'Requirements:
'
'<?>======================================================================================================================================================

Public Enum enuMultiply
    eAsNumber = 1
    eAsString = 0
End Enum

Public Enum enuDictionaryItemsType
    eValue = 1
    eArray = 2
    eRange = 3
End Enum
    
Public FSO As FileSystemObject
Public FD As FileDialog
Public WBK As Workbook
Public WSH As Worksheet
Public MsgTxt As String
Public ProcName As String
Public D As Dictionary
Public globHeader As Variant

Public rGPN As Range
Public rCountry As Range
Public rGCRSDivDesc As Range

Public Const HRI_REPORT_HEADER_ROW As Long = 22
Public Const HRI_REPORT_COLUMNS_COUNT As Long = 64
Public Const HRI_REPORT_ADDITIONAL_COLUMNS As Integer = 3
'

'?******************************************************************************************************************************************************
'Content Number:    1
'Name:              ClearTable
'Purpose:
'?******************************************************************************************************************************************************
Sub ClearTable(Ws As Worksheet, ListObjectName As String)
    Set loTable = Ws.ListObjects(ListObjectName)
    On Error Resume Next
    loTable.DataBodyRange.Delete
    On Error GoTo 0
End Sub
'?******************************************************************************************************************************************************
'Content Number:    2
'Name:              OpenWorkbook
'Purpose:
'?******************************************************************************************************************************************************
Function OpenWorkbook(ByVal WbkFullPath As String) As Workbook
    Dim WbkName As String
    
    Set FSO = New FileSystemObject
    WbkName = FSO.GetFileName(WbkFullPath)
On Error Resume Next
    Set OpenWorkbook = Workbooks(WbkName)
    If Err <> 0 Then
        Set OpenWorkbook = Workbooks.Open(Filename:=WbkFullPath, UpdateLinks:=False, ReadOnly:=True)
    End If
On Error GoTo 0
End Function
'?******************************************************************************************************************************************************
'Content Number:    3
'Name:              FindTextInRange
'Purpose:
'?******************************************************************************************************************************************************
Function FindTextInRange(sFindingTxt As String, SerchInrng As Range) As Range
   Set FindTextInRange = SerchInrng.Find( _
        What:=sFindingTxt, _
        LookAt:=xlPart, _
        LookIn:=xlValues, _
        MatchCase:=False _
        )
End Function
'?******************************************************************************************************************************************************
'Content Number:    4
'Name:              GetListObjectLastRowCell
'Purpose:
'?******************************************************************************************************************************************************
Function GetListObjectLastRowCell(loTable As ListObject) As Range
    
    If loTable.ListRows.Count = 0 Then
        Set Rng = loTable.HeaderRowRange.Cells(1).Offset(1)
    Else
        Set Rng = loTable.ListRows.item(loTable.ListRows.Count).Range.Cells(1)
    End If
    Debug.Print Rng.Address 'TODEL
    Set GetListObjectLastRowCell = Rng
End Function

'?******************************************************************************************************************************************************
'Content Number:    5
'Name:              PasteArray
'Purpose:
'?******************************************************************************************************************************************************
Sub PasteArray(ByRef vArr As Variant, ByVal FirstCell As Range, Optional sColFormat As String) ', Convert As enuMultiply)
    Dim DestRng As Range
    
    'Setting
    Set DestRng = FirstCell.Resize(UBound(vArr, 1), UBound(vArr, 2))
    
    'Formating
    If Not sColFormat = vbNullString Then DestRng.EntireColumn.NumberFormat = sColFormat
    
    'Pasting
    DestRng = vArr
    
End Sub

'?******************************************************************************************************************************************************
'Content Number:    7
'Name:              ExcludeUnsavedCharacters
'Purpose:
'?******************************************************************************************************************************************************
Function ExcludeUnsavedCharacters(ByVal sName As String) As String
    Dim lCharNo As Long
    Dim sChar As String
    Dim s As String
    
    ExcludeUnsavedCharacters = sName
    For lCharNo = 1 To Len(sName)
        sChar = Mid(sName, lCharNo, 1)
'        32-     space
'        48-57   digits
'        65-90   upper case
'        97-122  lower case
        If Not ( _
            Asc(sChar) = 32 _
            Or (Asc(sChar) >= 48 And Asc(sChar) <= 57) _
            Or (Asc(sChar) >= 65 And Asc(sChar) <= 90) _
            Or (Asc(sChar) >= 97 And Asc(sChar) <= 122) _
        ) Then
            ExcludeUnsavedCharacters = Replace(ExcludeUnsavedCharacters, sChar, vbNullString)
            
        End If
    Next
End Function

'?******************************************************************************************************************************************************
'Content Number:    8
'Name:              DictToArr
'Purpose:
'?******************************************************************************************************************************************************
Function CopyDictionaryToArray(ByVal objDic As Dictionary, eType As enuDictionaryItemsType) As Variant
    Dim T As Variant
    Dim lRowCount As Long, lRow As Long
    Dim lColCount As Long, lCol As Long
    Dim vKey As Variant
    Dim rCell As Range
    Dim V As Variant
    
    With objDic
    Select Case eType
        Case eValue
            lRowCount = .Keys.Count
            lColCount = .Items(1).Cells
            
            
        Case eArray
            lRowCount = .Keys.Count
            lColCount = .Items(1).Cells
            
            
        Case eRange
            lRowCount = .Count
            lColCount = .Items(0).Cells.Count
            
            ReDim T(1 To lRowCount, 1 To lColCount)
            
            For Each vKey In .Keys
                lRow = lRow + 1
                lCol = 0
                For Each rCell In objDic(vKey).Cells
                    lCol = lCol + 1
                    T(lRow, lCol) = rCell.Value
                Next rCell
            Next vKey
    End Select
    End With
    
    CopyDictionaryToArray = T
    
End Function
