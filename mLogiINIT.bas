Option Compare Database
Option Explicit

Private Const CurrentModuleName = "mLogiINIT"

Const LogON = True

Public Sub DopiszDoLogow(newData As String, Optional AddPomPar As Boolean = True, Optional EmptyLinePrefix As Boolean = False, Optional LogFilename As String = "")
    Dim LinePrefix As String
    Dim TopLine As String
    Dim pomPar As String
    Dim werIdPracownikaZalogowanego As Long

      Const sfName = "DopiszDoLogow"
      Dim ErrNumber, ErrDescription
On Error GoTo Err_PROCEDURE
10  werIdPracownikaZalogowanego = 0 ' TO DO - put  funkction return User ID
20  If LogON = False Then Exit Sub

30  pomPar = ""
40  If AddPomPar Then
50      If werIdPracownikaZalogowanego > 0 Then
60          pomPar = "PrId:" & werIdPracownikaZalogowanego & ","
70      Else
80          pomPar = "PrId:-,"
90      End If
100     pomPar = " (" & pomPar & "UN:" & GetUserName & ")"
110 End If

120 If SingleLogFile Then
130     LinePrefix = Now & pomPar
140 Else
150     LinePrefix = time & pomPar
160 End If

170 If EmptyLinePrefix Then
180     LinePrefix = String(Len(LinePrefix), " ")
190 End If

200 If Len(Trim(LinePrefix)) > 0 Then
210     TopLine = LinePrefix & " : " & newData
220 Else
230     TopLine = newData
240 End If

250 If Len(LogFilename) = 0 Then
260     Call InsertLineAtBeginningTexFile(TopLine)
270 ElseIf Len(LogFilename) > 3 Then
280     Call InsertLineAtBeginningTexFile(TopLine, LogFilename)
290 End If

Exit_PROCEDURE:
    Exit Sub

Err_PROCEDURE:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    Select Case ErrNumber

    Case Else
        VBA.MsgBox "(" & ErrNumber & IIf(Erl = 0, "", "," & Erl) & ") " & _
               "- " & CurrentModuleName & "." & sfName & vbLf & _
               ErrDescription, vbOKOnly + vbInformation, "Uwaga"
    End Select
    Resume Exit_PROCEDURE
    Resume

End Sub
