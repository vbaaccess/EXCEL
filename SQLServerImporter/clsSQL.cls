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
Private fu As New clsHelper

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
    
    Dim vD As Variant
    Dim sD$
    Call xls.Settings.ReadWhereVariable
    
    '-- 1 --
    vD = Null
    vD = xls.Settings.WhereCompany
    
    If Not IsNull(vD) Then
        If Len(Trim(CStr(vD))) > 0 Then
            sD = fu.KillSQLInjection(CStr(vD))
            SqlW = SqlW & " AND Kod = '" & CStr(sD) & "'"
        End If
    End If
    
    '-- 2 --
    vD = Null
    vD = xls.Settings.WhereFromDate
    
    If IsDate(vD) Then
        sD = fu.KillSQLInjection(CStr(vD))
        SqlW = SqlW & " AND OkresFrom >= convert(DATETIME,'" & sD & " 00:00:00.000',120)"
    End If
    
    '-- 3 --
    vD = Null
    vD = xls.Settings.WhereToDate
    
    If IsDate(vD) Then
        sD = fu.KillSQLInjection(CStr(vD))
        SqlW = SqlW & " AND OkresFrom <= convert(DATETIME,'" & sD & " 23:59:59.999',120)"
    End If
    
    If Len(SqlW) > 0 Then
        SqlW = "WHERE" & Mid(SqlW, 5)
    End If
    
    SQL_WHERE = SqlW
End Function

Private Function SQL_ORDER() As String
    SQL_ORDER = ""
End Function
