VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CurrentModuleName = "clsSQL"

Public Function GenerateSQL() As String
    Dim Sql As String
    Dim SqlS As String
    Dim SqlF As String
    Dim SqlW As String
    Dim SqlO As String
    
    Call RETURN_SQL(SqlS, SqlF, SqlW, SqlO)
    
    Sql = SqlS & SqlF & SqlW & SqlO
    
    GenerateSQL = Sql
End Function

Public Function RETURN_SQL(ByRef sSELECT As String _
                         , ByRef sFROM As String _
                         , ByRef sWHERE As String _
                         , ByRef sORDER As String) As Boolean

    sSELECT = SQL_SELECT()
    sFROM = SQL_FROM()
    sWHERE = SQL_WHERE()
    sORDER = SQL_ORDER()
    
    RETURN_SQL = True
End Function

Private Function SQL_SELECT() As String
    Dim Sql As String
    Sql = "SELECT @SQL_TOP_SQL * "
    
    Dim iDataSampleSize As Long
    Dim sDataSampleSize As String
    
    iDataSampleSize = xls.Settings.DataSampleSize
    If iDataSampleSize > 0 Then
        sDataSampleSize = "TOP " & iDataSampleSize
    End If
    
    Sql = Replace(Sql, "@SQL_TOP_SQL", sDataSampleSize)
    SQL_SELECT = Sql
End Function

Private Function SQL_FROM() As String
    SQL_FROM = " FROM vForXLS "
End Function

Private Function SQL_WHERE() As String
    
    Dim SqlW As String
    Dim d As Dictionary
    
    
    Set d = xls.Settings.DataConditions.ReadTheData()
    
    Dim v
    
    For Each v In d
        SqlW = SqlW & d(CStr(v))
    Next v
    
    If Len(SqlW) > 0 Then
        SqlW = "WHERE" & Mid(SqlW, 5)
        Debug.Print SqlW
    End If
    
    SQL_WHERE = SqlW
End Function

Private Function SQL_ORDER() As String
    SQL_ORDER = ""
End Function
