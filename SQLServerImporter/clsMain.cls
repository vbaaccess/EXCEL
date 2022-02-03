Option Explicit

Private Const CurrentModuleName = "clsMain"
Public Settings As clsSettings

Private Sub Class_Initialize()
    'Constructor
    Debug.Print "Initialize object clas: " & CurrentModuleName
    
    If Settings Is Nothing Then Set Settings = New clsSettings
    Set Settings = Settings
End Sub

Private Sub Class_Terminate()
    'Destruct
    Debug.Print "Terminate object clas: " & CurrentModuleName
End Sub

Public Sub ImportData()
    Dim ModuleName, SubName
    Dim ErrNumber, ErrDescription
On Error GoTo Err_SUB

    Call ImportProcedures
    
Exit_SUB:
    Exit Sub
    
Err_SUB:
    SubName = "ImportData"
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    Select Case ErrNumber
    Case Else
        VBA.MsgBox "(" & ErrNumber & IIf(Erl = 0, "", "," & Erl) & ") " & _
               "- " & CurrentModuleName & "." & SubName & vbLf & _
               ErrDescription, vbOKOnly + vbInformation, "Uwaga"
    End Select
    Resume Exit_SUB
    Resume
End Sub

Private Sub ImportProcedures()
    Dim Rst As Recordset
    Dim Sql As String
    

    Sql = GenerateSQL
    
    If OpenRst(Rst, Sql) Then
        Debug.Print "Rst.RecordCount => " & CStr(Rst.RecordCount)
        
        Rst.MoveLast
        
        Call ImportRecordsHeaders(Rst)
        Call ImportRecordsData(Rst)
        
    End If
    
    Call rstSetNothing(Rst)
End Sub

Private Function GenerateSQL() As String
    Dim Sql As String
    Dim SqlW As String
    Dim iDataSampleSize As Long
    Dim sDataSampleSize As String
    
    iDataSampleSize = xls.Settings.DataSampleSize
    If iDataSampleSize > 0 Then
        sDataSampleSize = "TOP " & iDataSampleSize
    End If
    
    Dim vD As Variant
    Call xls.Settings.ReadWhereVariable
    
    '-- 1 --
    vD = Null
    vD = xls.Settings.WhereFromDate
    
    If IsDate(vD) Then
        SqlW = SqlW & " AND "
    End If
    
    '-- 2 --
    vD = Null
    vD = xls.Settings.WhereToDate
    
    If IsDate(vD) Then
        SqlW = SqlW & " AND "
    End If
    
    If Len(SqlW) > 0 Then
        SqlW = "WHERE" & Mid(SqlW, 4)
    End If
    
    Sql = "SELECT @SQL_TOP_SQL * FROM vForXLS @WhereCondytion"
    
    Sql = Replace(Sql, "@SQL_TOP_SQL", sDataSampleSize)
    Sql = Replace(Sql, "@WhereCondytion", SqlW)
    
    GenerateSQL = Sql
End Function

Private Function SpecifyTheSheet(Rst As Recordset) As Worksheet
    Dim wb As Workbook
    Dim wks As Worksheet
    
End Function

Private Sub ImportRecordsHeaders(Rst As Recordset)
        
        Rst.MoveFirst
        Dim fld
        Dim SheetName$
        SheetName = xls.Settings.TargetSheets
            
        Dim cLp As Long
        Dim rLp As Long
        Dim posX As Long
        Dim posY As Long
        
        Dim shift_x As Long
        Dim shift_y As Long
        
        shift_x = xls.Settings.SheetsPositionsX
        shift_y = xls.Settings.SheetsPositionsY
        If shift_x > 0 Then shift_x = shift_x - 1
        If shift_y > 0 Then shift_y = shift_y - 1
        
        rLp = 1
        
        
        
        For cLp = 0 To Rst.Fields.Count - 1
            posX = cLp + 1 + shift_x
            posY = rLp + shift_y
            
            
            Sheets(SheetName).Cells(posY, posX) = Rst.Fields(cLp).Name
            Sheets(SheetName).Cells(posY, posX).Font.Bold = True
            
            ThisWorkbook.Sheets(SheetName).Cells(posY, posX) = Rst.Fields(cLp).Name
            ThisWorkbook.Sheets(SheetName).Cells(posY, posX).Font.Bold = True
        Next cLp
        
        Do While Not Rst.EOF
            rLp = rLp + 1
            cLp = 0
            For Each fld In Rst.Fields
                cLp = cLp + 1
                posX = cLp + shift_x
                posY = rLp + shift_y
                ThisWorkbook.Sheets(SheetName).Cells(posY, posX) = fld
                If IsDate(fld) Then
                    ThisWorkbook.Sheets(SheetName).Cells(posY, posX).NumberFormat = "m/d/yyyy"
                End If
            Next
            Rst.MoveNext
        Loop
End Sub

Private Sub ImportRecordsData(Rst As Recordset)
    Rst.MoveFirst
End Sub


Private Function ActiveConnection() As Connection
    Dim sConnectionString$
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    Dim cs As New clsConnectionString
    
    sConnectionString = cs.ConnectionString
        
    conn.ConnectionString = sConnectionString
    conn.Open
    
    Set ActiveConnection = conn
End Function

Private Function OpenRst(ByRef Rst As ADODB.Recordset, ByVal Sql As String) As Boolean   
    Dim conn As ADODB.Connection
    
    Set conn = ActiveConnection
    
    Set Rst = New ADODB.Recordset

    Rst.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdUnknown

    If Rst.RecordCount <= 0 Then
        OpenRst = False
    Else
        OpenRst = True
    End If
End Function

Public Sub rstSetNothing(ByRef Rst As ADODB.Recordset)
On Error GoTo rstSetNothing_Error

    If Rst Is Nothing Then Exit Sub

    If Rst.State = adStateOpen Then
        Rst.Close
    End If

    Set Rst = Nothing

    On Error GoTo 0
    Exit Sub

rstSetNothing_Error:
    Exit Sub
End Sub
