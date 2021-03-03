Option Compare Database
Option Explicit

Private Const CurrentModuleName = "mFunkcjeRecordset"

Public Function OpenRst(ByRef Rst As ADODB.Recordset, ByVal Sql As String , Optional SqlTimeout As Long = 0 ) As Boolean
    Set Rst = New ADODB.Recordset

    If SqlTimeout > 0 Then
        Dim con As ADODB.Connection
        Set con = CurrentProject.Connection
        con.CommandTimeout = SqlTimeout
        
        Rst.Open Sql, con, adOpenKeyset, adLockOptimistic, adCmdUnknown
    Else
        Rst.Open Sql, CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdUnknown
    End If

    If Rst.RecordCount <= 0 Then
        OpenRst = False
    Else
        OpenRst = True
    End If

End Function
