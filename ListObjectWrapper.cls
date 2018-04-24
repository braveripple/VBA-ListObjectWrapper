VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListObjectWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private list As ListObject

''' ListObjectWrapper Property Set/Let

Property Set ListObject(ByVal obj As ListObject)
    Set list = obj
End Property

''' ListObjectWrapper Property Get

Property Get ListObject() As ListObject
    Set ListObject = list
End Property

Property Get Worksheet() As Worksheet
    Set Worksheet = list.Parent
End Property

Property Get Workbook() As Workbook
    Set Workbook = list.Parent.Parent
End Property

Property Get TableName() As String
    TableName = "[" & Me.Worksheet.Name & "$" & Me.Range.Address(False, False) & "]"
End Property

Property Get TableNameStrict() As String
    If (Me.Range.Cells(1, 1).Address = "$A$1") Then
        ' [SheetName$]
        TableNameStrict = "[" & Me.Worksheet.Name & "$" & "]"
    Else
        ' [SheetName$RangeName(A1)]
        TableNameStrict = "[" & Me.Worksheet.Name & "$" & Me.Range.Address(False, False) & "]"
    End If
End Property

Property Get FirstRow() As Range
    Set FirstRow = list.ListRows(1).Range
End Property

Property Get FirstRowFirstCell() As Range
    Set FirstRowFirstCell = Me.FirstRow.Cells(1, 1)
End Property

Property Get FirstRowLastCell() As Range
    Set FirstRowLastCell = Me.FirstRow.Cells(1, list.ListColumns.Count)
End Property

Property Get FirstRowColumnCell(ByVal Index As Variant) As Range
    Set FirstRowColumnCell = list.ListColumns(Index).DataBodyRange(1, 1)
End Property


Property Get LastRow() As Range
    Set LastRow = list.ListRows(list.ListRows.Count).Range
End Property

Property Get LastRowFirstCell() As Range
    Set LastRowFirstCell = Me.LastRow.Cells(1, 1)
End Property

Property Get LastRowLastCell() As Range
    Set LastRowLastCell = Me.LastRow.Cells(1, list.ListColumns.Count)
End Property

Property Get LastRowColumnCell(ByVal Index As Variant) As Range
    Set LastRowColumnCell = list.ListColumns(Index).DataBodyRange(list.ListRows(1).Count, 1)
End Property


Property Get ColumnRow() As Range
    Set ColumnRow = list.HeaderRowRange
End Property

Property Get ColumnRowFirstCell() As Range
    Set ColumnRowFirstCell = Me.ColumnRow.Cells(1, 1)
End Property

Property Get ColumnRowLastCell() As Range
    Set ColumnRowLastCell = Me.ColumnRow.Cells(1, list.ListColumns.Count)
End Property

Property Get ColumnRowColumnCell(ByVal Index As Variant) As Range
    Set ColumnRowColumnCell = list.ListColumns(Index).Range
End Property


''' ListObjectWrapper Public Method

Public Sub ResetFilter(ByVal showAllData As Boolean)
    If list.AutoFilter Is Nothing Then
        ' Set Filter
        list.Range.AutoFilter
    Else
        ' If filter is set, all data is displayed
        If showAllData Then
            list.AutoFilter.showAllData
        End If
    End If
End Sub

Public Sub ClearData()
    
    Call ResetFilter(showAllData:=True)
    
    With list
        On Error Resume Next
        .DataBodyRange.Delete xlShiftUp
        On Error GoTo 0
        Me.AddRow
        Me.Worksheet.Rows(Me.LastRow.Row + 1 & ":" & Me.Worksheet.Rows.Count).EntireRow.AutoFit
    End With

End Sub

Public Sub AddRow()
    
    With list
        ' Create data row
        If .ListRows.Count = 0 Then
            ' If the number of data rows is 0, data pasting will cause a runtime error,
            ' so add data rows in advance.
            .ListRows.Add
        Else
            If .ListRows(.ListRows.Count).Range(1, 1) = "" Then
                ' There is no need to add data rows
                Exit Sub
            End If
            .ListRows.Add
        End If
    End With

End Sub

Public Sub CopyFromRecordset(ByVal Data, Optional ByVal Append = False)
    
    With Me
        If Data.EOF Then
            If Append = False Then
                .ClearData
            End If
            Exit Sub
        End If
        
        If Append Then
            .AddRow
        Else
            .ClearData
        End If
        
        .FirstRowFirstCell.CopyFromRecordset Data
    End With
    
End Sub

Public Sub ResizeRow(ByVal RowSize As Long)
    Call list.Resize(list.Range.Resize(RowSize + 1))
End Sub

Public Function IsEmpty() As Boolean
    With list
        If (.ListRows.Count = 0) Then
            IsEmpty = True
            Exit Function
        ElseIf (.ListRows.Count = 1 And .ListColumns(1).DataBodyRange(1).Value = "") Then
            IsEmpty = True
            Exit Function
        End If
    End With
    IsEmpty = False
End Function

''' ListObject Wrapper Properties.

Public Property Get ListColumns() As ListColumns
    Set ListColumns = list.ListColumns
End Property

Property Get ListRows() As ListRows
    Set ListRows = list.ListRows
End Property

Property Get Name() As String
    Name = list.Name
End Property

Property Get DisplayName() As String
    DisplayName = list.DisplayName
End Property

Property Get Range() As Range
    Set Range = list.Range
End Property

Property Get DataBodyRange() As Range
    Set DataBodyRange = list.DataBodyRange
End Property

Property Get HeaderRowRange() As Range
    Set HeaderRowRange = list.HeaderRowRange
End Property

Property Get InsertRowRange() As Range
    Set InsertRowRange = list.InsertRowRange
End Property

Property Get TotalsRowRange() As Range
    Set TotalsRowRange = list.TotalsRowRange
End Property

Property Get SourceType() As XlListObjectSourceType
    SourceType = list.SourceType
End Property

Property Get Parent() As Worksheet
    Set Parent = list.Parent
End Property

Property Get Application() As Application
    Set Application = list.Application
End Property

Property Get ShowAutoFilter() As Boolean
    ShowAutoFilter = list.ShowAutoFilter
End Property

Property Let ShowAutoFilter(ByVal val As Boolean)
    list.ShowAutoFilter = val
End Property

Property Get Sort() As Sort
    Set Sort = list.Sort
End Property

''' ListObject Wrapper Methods.

Sub Delete()
    list.Delete
End Sub

Sub Refresh()
    list.Refresh
End Sub

Sub Resize(ByVal Range As Range)
    list.Resize Range
End Sub

Sub Unlink()
    list.Unlink
End Sub

Sub Unlist()
    list.Unlist
End Sub
