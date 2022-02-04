VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CurrentModuleName = "clsMain"
Public Settings As clsSettings
Attribute Settings.VB_VarHelpID = -1

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
'    Case NrBledu
'        Debug.Print "(" & ErrNumber  & ") - " & ModuleName & "." & SubName

    Case Else
        VBA.MsgBox "(" & ErrNumber & IIf(Erl = 0, "", "," & Erl) & ") " & _
               "- " & CurrentModuleName & "." & SubName & vbLf & _
               ErrDescription, vbOKOnly + vbInformation, "Uwaga"
    End Select
    Resume Exit_SUB
    Resume
End Sub

Private Sub ImportProcedures()
    Dim wks As Worksheet
    Dim Rst As Recordset
    Dim Sql As String
    
    Dim SqlG As New clsSQL

    Sql = SqlG.GenerateSQL()
    
    If OpenRst(Rst, Sql) Then
        Debug.Print "Rst.RecordCount => " & CStr(Rst.RecordCount)

        Set wks = SpecifyTheSheet()
        
        Call ImportRecordsHeaders(Rst, wks)
        Call ImportRecordsData(Rst, wks)
        
    End If
    
    Call rstSetNothing(Rst)
End Sub

Private Function SpecifyTheSheet() As Worksheet
    Dim f$, s$  ' file name, sheet
    Dim wb As Workbook
    Dim wks As Worksheet
    
    '-- 1 --- set Workbook
    f = xls.Settings.TargetFileName
    If Len(Trim(f)) > 0 Then
        Set wb = SetWorkbooks(f)
    End If
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    '-- 2 --- set WorkSheet
    s = xls.Settings.TargetSheets
    Set wks = wb.Worksheets(s)
    
    Set SpecifyTheSheet = wks
End Function

Private Function SetWorkbooks(FileName As String) As Workbook
On Error Resume Next
    Set SetWorkbooks = Workbooks(FileName)
End Function

Private Sub InitShift(x As Long, y As Long)
    x = xls.Settings.SheetsPositionsX
    y = xls.Settings.SheetsPositionsY
    If x > 0 Then x = x - 1
    If y > 0 Then y = y - 1
End Sub

Private Sub ImportRecordsHeaders(Rst As Recordset, wks As Worksheet)
    Dim cLp As Long
    
    Dim shift_x As Long
    Dim shift_y As Long
    
    Call InitShift(shift_x, shift_y)
        
    Dim posX As Long
    Dim posY As Long
        
    posY = 1 + shift_y
    
    For cLp = 0 To Rst.Fields.Count - 1
        posX = cLp + 1 + shift_x
        wks.Cells(posY, posX) = Rst.Fields(cLp).Name
        wks.Cells(posY, posX).Font.Bold = True
    Next cLp

End Sub

Private Sub ImportRecordsData(Rst As Recordset, wks As Worksheet)
    Dim cLp As Long
    Dim fld
            
    Dim shift_x As Long
    Dim shift_y As Long
    
    Call InitShift(shift_x, shift_y)
        
    Dim posX As Long
    Dim posY As Long
    
    posY = 1 + shift_y
    posX = shift_x
    Rst.MoveFirst
        
    Do While Not Rst.EOF
        posY = posY + 1
        posX = shift_x

        For Each fld In Rst.Fields
            posX = posX + 1
            wks.Cells(posY, posX) = fld
            If IsDate(fld) Then
                wks.Cells(posY, posX).NumberFormat = "m/d/yyyy"
            End If
        Next
        Rst.MoveNext
    Loop

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
