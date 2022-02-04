VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSetings 
   Caption         =   "Settings"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "UserFormSetings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSetings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBoxAuthentication_Change()
    Call AfterChange
End Sub

Private Sub CommandButtonBack_Click()
    Unload Me
End Sub

Private Sub CommandButtonCancel_Click()
    Call InitializeAllSettings
End Sub

Private Sub CommandButtonSave_Click()
    Call SaveProcedure

    Call Me.CommandButtonBack.SetFocus
    Me.CommandButtonCancel.Enabled = False
    Me.CommandButtonSave.Enabled = False
End Sub

Private Sub AfterChange()
    Me.CommandButtonCancel.Enabled = True
    Me.CommandButtonSave.Enabled = True
End Sub

Private Sub TextBoxDataSampleSize_Change()
    Call AfterChange
End Sub

Private Sub TextBoxDBName_Change()
    Call AfterChange
End Sub

Private Sub TextBoxLogin_Change()
    Call AfterChange
End Sub

Private Sub TextBoxPositionX_Change()
    Call AfterChange
End Sub

Private Sub TextBoxPositionY_Change()
    Call AfterChange
End Sub

Private Sub TextBoxPWD_Change()
    Call AfterChange
End Sub

Private Sub TextBoxServerName_Change()
    Call AfterChange
End Sub

Private Sub TextBoxTargetFile_Change()
    Call AfterChange
End Sub

Private Sub TextBoxTargetSheets_Change()
    Call AfterChange
End Sub

Private Sub UserForm_Initialize()
    Call InitializeAllSettings
End Sub

Private Sub InitializeAllSettings()
    Call InitPage1_XLS
    Call InitPage2_DB
    Call InitPage3_Import
    
    Call Me.CommandButtonBack.SetFocus
    Me.CommandButtonCancel.Enabled = False
    Me.CommandButtonSave.Enabled = False

End Sub

Private Sub InitPage1_XLS()
    Me.TextBoxTargetFile = xls.Settings.TargetFileName
    Me.TextBoxTargetSheets = xls.Settings.TargetSheets
    
    Me.TextBoxPositionX = xls.Settings.SheetsPositionsX
    Me.TextBoxPositionY = xls.Settings.SheetsPositionsY
End Sub

Private Sub InitPage2_DB()
    Dim sServerType$
    sServerType = "Microsoft SQL Server"
    With Me.ComboBoxServerType
        .AddItem sServerType
        .Value = sServerType
        .Enabled = False
    End With
    
    Me.TextBoxServerName = xls.Settings.Server
    Me.TextBoxDBName = xls.Settings.InitialCatalog
        
    Call AuthenticationInitialize
    
    Me.TextBoxLogin = xls.Settings.Login
    Me.TextBoxPWD = xls.Settings.Password
    
    Me.TextBoxDBName = xls.Settings.InitialCatalog

End Sub

Private Sub InitPage3_Import()
    Me.TextBoxDataSampleSize = xls.Settings.DataSampleSize
End Sub


Private Sub AuthenticationInitialize()
    Dim sAuthentication(0 To 1) As String
    sAuthentication(0) = "WindowsAuthentication"
    sAuthentication(1) = "Sql Server Authentication"
    
    With Me.ComboBoxAuthentication
        .AddItem sAuthentication(0)
        .AddItem sAuthentication(1)
        .Value = sAuthentication(xls.Settings.Authentication)
        .Enabled = True
    End With
End Sub

Private Sub AuthenticationSave()
    Dim sAuthentication(0 To 1) As String
    sAuthentication(0) = "WindowsAuthentication"
    sAuthentication(1) = "Sql Server Authentication"
    
    With Me.ComboBoxAuthentication
        If .Value = sAuthentication(0) Then
            xls.Settings.Authentication = 0
        End If
        If .Value = sAuthentication(1) Then
            xls.Settings.Authentication = 1
        End If
    End With
End Sub

Private Sub SaveProcedure()
    'Page 1
    xls.Settings.TargetFileName = Me.TextBoxTargetFile
    xls.Settings.TargetSheets = Me.TextBoxTargetSheets
    xls.Settings.SheetsPositionsX = Me.TextBoxPositionX
    xls.Settings.SheetsPositionsY = Me.TextBoxPositionY
    
    'Page 2
    xls.Settings.Server = Me.TextBoxServerName
    xls.Settings.InitialCatalog = Me.TextBoxDBName
    
    xls.Settings.Login = Me.TextBoxLogin
    xls.Settings.Password = Me.TextBoxPWD
    
    Call AuthenticationSave
    
    xls.Settings.InitialCatalog = Me.TextBoxDBName
    
    'Page 3
    xls.Settings.DataSampleSize = Me.TextBoxDataSampleSize
End Sub
