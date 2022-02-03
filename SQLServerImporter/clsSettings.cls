Option Explicit

Private Const CurrentModuleName = "clsSettings"
Private Const m_SettingSheetsName = "XLS_Settings_XLS"
Private Const m_InstructionSheetsName = "Instrukcja"

Private Const m_CfgColumn = 2

'Public Type
Private Enum Cfg
    'Authentication = 2
    Auth = 2
    UserName = 3
    PWD = 4
    'Server = 6
    Svr = 6
    'InitialCatalog = 7
    dbName = 7
    'TargetSheets = 9
    SettingsSheets = 9
    SettingsSheetsPositionX = 10
    SettingsSheetsPositionY = 11
    SampleSize = 14
End Enum

Private m_Authentication As Single
Private m_UserName As String
Private m_PWD As String

Private m_Server As String
Private m_InitialCatalog As String

Private m_TargetSheets As String
Private m_TargetPositionX As Long
Private m_TargetPositionY As Long

Private m_DataSampleSize As Long

Private m_FromDate As Variant
Private m_ToDate As Variant

Private Sub Class_Initialize()
    Debug.Print "Initialize object clas: " & CurrentModuleName
    
    m_FromDate = Null
    m_ToDate = Null
    
    Call RefreshSettings

End Sub

Public Sub RefreshSettings()
        'xls
        m_TargetSheets = GetSetting(SettingsSheets)
        m_TargetPositionX = GetSetting(SettingsSheetsPositionX)
        m_TargetPositionY = GetSetting(SettingsSheetsPositionY)
        
        'db
        m_Authentication = GetSetting(Auth)
        m_UserName = GetSetting(UserName)
        m_PWD = GetSetting(PWD)
        
        m_Server = GetSetting(Svr)
        m_InitialCatalog = GetSetting(dbName)
        
        'import
        m_DataSampleSize = GetSetting(SampleSize)
        Call ReadWhereVariable
End Sub

Public Sub ReadWhereVariable()
    Dim vD As Variant
    
    With ThisWorkbook.Sheets(m_InstructionSheetsName)
        vD = Null
        vD = .Cells(10, m_CfgColumn).Value
        If IsDate(vD) Then
            m_FromDate = vD
        End If
        
        vD = Null
        vD = .Cells(11, m_CfgColumn).Value
        If IsDate(vD) Then
            m_ToDate = vD
        End If
        
    End With
End Sub

Private Function GetSetting(cell As Cfg) As Variant
    With ThisWorkbook.Sheets(m_SettingSheetsName)
        GetSetting = .Cells(cell, m_CfgColumn).Value
    End With
End Function

Private Sub SetSetting(cell As Cfg, NewValue As Variant)
    With ThisWorkbook.Sheets(m_SettingSheetsName)
        .Cells(cell, m_CfgColumn).Value = NewValue
    End With
End Sub

Public Sub Show()
    Dim ModuleName, SubName
    Dim ErrNumber, ErrDescription
On Error GoTo Err_SUB

    UserFormSetings.Show
    
Exit_SUB:
    Exit Sub
    
Err_SUB:
    SubName = "Show"
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

Public Property Get Authentication() As Single
    Authentication = m_Authentication
End Property

Public Property Let Authentication(Authentication_No As Single)
    If Authentication_No <> m_Authentication Then _
        Call SetSetting(Auth, Authentication_No)

    m_Authentication = Authentication_No
End Property

Public Property Get Login() As String
    Login = m_UserName
End Property

Public Property Let Login(NewUserName As String)
    If NewUserName <> m_UserName Then _
        Call SetSetting(UserName, NewUserName)
        
    m_UserName = NewUserName
End Property

Public Property Get Password() As String
    Password = m_PWD
End Property

Public Property Let Password(NewPassword As String)
    If NewPassword <> m_PWD Then _
        Call SetSetting(PWD, NewPassword)
        
    m_PWD = NewPassword
End Property

Public Property Get Server() As String
    Server = m_Server
End Property

Public Property Let Server(NewServer As String)
    If NewServer <> m_Server Then _
        Call SetSetting(Svr, NewServer)
        
    m_Server = NewServer
End Property

Public Property Get InitialCatalog() As String
    InitialCatalog = m_InitialCatalog
End Property

Public Property Let InitialCatalog(NewInitialCatalog As String)
    If NewInitialCatalog <> m_InitialCatalog Then _
        Call SetSetting(dbName, NewInitialCatalog)
        
    m_InitialCatalog = NewInitialCatalog
End Property

Public Property Get TargetSheets() As String
    TargetSheets = m_TargetSheets
End Property

Public Property Let TargetSheets(NewTargetSheets As String)
    If NewTargetSheets <> m_TargetSheets Then _
        Call SetSetting(SettingsSheets, NewTargetSheets)
        
    m_TargetSheets = NewTargetSheets
End Property

Public Property Get SheetsPositionsX() As String
    SheetsPositionsX = m_TargetPositionX
End Property

Public Property Let SheetsPositionsX(NewPositionX As String)
    If NewPositionX <> m_TargetPositionX Then _
        Call SetSetting(SettingsSheetsPositionX, NewPositionX)
        
    m_TargetPositionX = NewPositionX
End Property

Public Property Get SheetsPositionsY() As String
    SheetsPositionsY = m_TargetPositionY
End Property

Public Property Let SheetsPositionsY(NewPositionY As String)
    If NewPositionY <> m_TargetPositionY Then _
        Call SetSetting(SettingsSheetsPositionY, NewPositionY)
        
    m_TargetPositionY = NewPositionY
End Property

Public Property Get DataSampleSize() As Long
    DataSampleSize = m_DataSampleSize
End Property

Public Property Let DataSampleSize(NewSampleSize As Long)
    If NewSampleSize <> m_DataSampleSize Then _
        Call SetSetting(SampleSize, NewSampleSize)
        
    m_DataSampleSize = NewSampleSize
End Property

Public Property Get WhereFromDate() As Variant
    WhereFromDate = m_FromDate
End Property

Public Property Get WhereToDate() As Variant
    WhereToDate = m_ToDate
End Property
