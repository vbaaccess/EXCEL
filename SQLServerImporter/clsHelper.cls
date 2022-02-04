VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CurrentModuleName = "clsHelper"

Public Function KillSQLInjection(ByVal SqlString As String) As String
    Dim CorrectSQLString As String

    If Len(SqlString) = 0 Then
        CorrectSQLString = SqlString
    Else
        CorrectSQLString = Replace(SqlString, "'", "''")
    End If

    KillSQLInjection = CorrectSQLString
End Function

