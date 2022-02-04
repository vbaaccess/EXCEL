VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMsgBox 
   Caption         =   "Uwaga"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CurrentModuleName = "UserFormMsgBox"

Private Sub UserForm_Initialize()
    Me.CmdBt01.Caption = "TAK"
    Me.CmdBt03.Caption = "OK"
    Me.CmdBt03.Visible = False
    Me.CmdBt02.Caption = "NIE"
End Sub

Private Sub CmdBt01_Click()
    Call ProcedureYES
End Sub

Private Sub CmdBt03_Click()
    Call ProcedureNO
End Sub

Private Sub ProcedureYES()
    Debug.Print "UNDER CONSTRUCTION"
End Sub


Private Sub ProcedureNO()
    Debug.Print "UNDER CONSTRUCTION"
End Sub
