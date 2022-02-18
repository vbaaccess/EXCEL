VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataConditions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CurrentModuleName = "clsDataConditions"

Private Const m_ConditionsSheetsName = "Instrukcja"
Private Const m_DefalConditionsColumn = 2

Private fu As New clsHelper

Private dictConditions As New Dictionary

Private Sub InitConditions()
    Set dictConditions = Nothing
    dictConditions.Add "contract_beginning", ConditionSql("VARCHAR", 10)
    dictConditions.Add "contract_termination", ConditionSql("DATETIME", 15)
    dictConditions.Add "company_name", ConditionSql("DATETIME", 20)
    
End Sub

Private Function ConditionSql(DataType As String _
                                , row As Long _
                                , Optional column As Long _
                                , Optional SheetName As String _
                                ) As String
    Dim Sql As String ' Condition Sql String
    
    If column = 0 Then column = m_DefalConditionsColumn
    If Len(Trim(SheetName)) = 0 Then SheetName = m_ConditionsSheetsName
    
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets(SheetName)
    
    If DataType = "DATETIME" Then
        ConditionSql = ConditionSqlData(sheet, row, column)
    End If
    
    If DataType = "VARCHAR" Then
        ConditionSql = ConditionSqlString(sheet, row, column)
    End If

End Function

Private Function ConditionSqlData(sheet As Worksheet, row As Long, column As Long) As String
    Dim Sql$
    
    Dim vFrom As Variant
    Dim vTo As Variant
    Dim sD$
    
    vFrom = Null
    vTo = Null
    
    Dim FieldName As String
    FieldName = sheet.Cells(row, column).Value
    FieldName = fu.KillSQLInjection(FieldName)
    
    If IsDate(sheet.Cells(row, column + 1).Value) Then
        vFrom = sheet.Cells(row, column + 1).Value
        sD = fu.KillSQLInjection(FormatDateTime(CDate(vFrom), vbShortDate))
        sD = Year(CDate(Now)) & "-" & Right("0" & Month(CDate(Now)), 2) & "-" & Right("0" & Day(CDate(Now)), 2)
        Sql = Sql & " AND " & FieldName & " >= convert(DATETIME,'" & sD & " 00:00:00.000',120)"
    End If
    
    If IsDate(sheet.Cells(row, column + 2).Value) Then
        vTo = sheet.Cells(row, column + 2).Value
        sD = fu.KillSQLInjection(FormatDateTime(CDate(vTo), vbShortDate))
        sD = Year(CDate(Now)) & "-" & Right("0" & Month(CDate(Now)), 2) & "-" & Right("0" & Day(CDate(Now)), 2)
        Sql = Sql & " AND (" & FieldName & " <= convert(DATETIME,'" & sD & " 23:59:59.999',120)"
        Sql = Sql & " OR " & FieldName & " Is Null)"
    End If
    
    ConditionSqlData = Sql
End Function

Private Function ConditionSqlString(sheet As Worksheet, row As Long, column As Long) As String
    Dim FieldName As String
    FieldName = sheet.Cells(row, column).Value
    FieldName = fu.KillSQLInjection(FieldName)
    
    Dim Sql$
    Dim vStr As String
    
    vStr = ""
    If Len(sheet.Cells(row, column + 1).Value) Then
        vStr = sheet.Cells(row, column + 1).Value
        vStr = fu.KillSQLInjection(CStr(vStr))
        Sql = Sql & ",'" & vStr & "'"
    End If
    
    vStr = ""
    If Len(sheet.Cells(row, column + 2).Value) Then
        vStr = sheet.Cells(row, column + 2).Value
        vStr = fu.KillSQLInjection(CStr(vStr))
        Sql = Sql & ",'" & vStr & "'"
    End If
    
    
    If Len(Trim(Sql)) > 0 Then
        Sql = " AND " & FieldName & " IN (" & Mid(Sql, 2) & ")"
    End If
    
    
    ConditionSqlString = Sql
End Function

Private Sub Class_Initialize()
    'Constructor
    Debug.Print "Initialize object clas: " & CurrentModuleName
End Sub

Public Function ReadTheData() As Dictionary
    'return dictionary with sql conditions string
    Call InitConditions
    Set ReadTheData = dictConditions
End Function
