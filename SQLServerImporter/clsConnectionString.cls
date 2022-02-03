Option Explicit

Private Const CurrentModuleName = "clsConnectionString"

Private m_ConnectionString As String
Private m_Server As String
Private m_Instancja As String
Private m_InitialCatalog As String

Private Sub Class_Initialize()
    Debug.Print "Initialize object clas: " & CurrentModuleName
    
    With ThisWorkbook.Sheets(1)
        m_Server = .Cells(6, 2).Value
        m_Instancja = .Cells(7, 2).Value
        m_InitialCatalog = .Cells(8, 2).Value
    End With
    
    Call InitConnectionString
    
End Sub

Private Sub InitConnectionString()
    Dim iAuthentication As Single
    Dim s$
    
    'default: Windows authentication on test db
    s = "Provider=SQLOLEDB.1;Auto Translate=True;Data Source=SQLServer;Initial Catalog=XLS_TEST;Integrated Security=SSPI;Persist Security Info=FALSE;Use Procedure for Prepare=1"
    
    iAuthentication = xls.Settings.Authentication
    
    
    If iAuthentication = 0 Then
        'Windows authentication
        s = "Provider=SQLOLEDB.1;Auto Translate=True;Data Source={m_Server};Initial Catalog={m_InitialCatalog};Integrated Security=SSPI;Persist Security Info=FALSE;Use Procedure for Prepare=1"
    End If
    
    If iAuthentication = 1 Then
        'SQL authentication
        s = "Provider=SQLOLEDB.1;Auto Translate=True;Data Source={m_Server};Initial Catalog={m_InitialCatalog};Integrated Security=SSPI;Persist Security Info=FALSE;Use Procedure for Prepare=1"
        s = "Provider=sqloledb;Data Source={m_Server};Initial Catalog={m_InitialCatalog};Password={m_pwd};Persist Security Info=TRUE;User Id={m_login}"
    End If
    
    
    s = Replace(s, "{m_Server}", xls.Settings.Server)
    s = Replace(s, "{m_InitialCatalog}", xls.Settings.InitialCatalog)
    s = Replace(s, "{m_login}", xls.Settings.Login)
    s = Replace(s, "{m_pwd}", xls.Settings.Password)
    
    m_ConnectionString = s
End Sub

Public Function ConnectionString() As String
    ConnectionString = m_ConnectionString
End Function
