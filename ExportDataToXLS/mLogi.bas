Option Compare Database
Option Explicit

Private Const CurrentModuleName = "mLogi"
Private Const PrefixLogFile = "sur_"
Private Const SufixLogFile = ".log"

Public Const SingleLogFile = True

Public Sub InsertLineAtBeginningTexFile(strLine As String, Optional LogFilename As String = "")
    Dim strFile As String
    Dim hFile As Long
    Dim FileContents As String
    Dim NewString As String

      Const sfName = "InsertLineAtBeginningTexFile"
      Dim ErrNumber, ErrDescription
On Error GoTo Err_PROCEDURE

20  strFile = Application.CurrentProject.Path
30  If Len(LogFilename) = 0 Then
40      strFile = strFile & "\" & NazwaPlikLogu
50  Else
60      strFile = strFile & "\" & LogFilename
70  End If

80  hFile = FreeFile
90  Open strFile For Binary As #hFile
100 FileContents = Space(FileLen(strFile))
110 Get #hFile, , FileContents
120 Close #hFile

130 NewString = strLine & vbCrLf
140 FileContents = NewString & FileContents

150 Open strFile For Binary As #hFile
160 Put #hFile, , FileContents
170 Close #hFile

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

Public Function NazwaPlikLogu() As String

      Const sfName = "NazwaPlikLogu"
      Dim ErrNumber, ErrDescription
On Error GoTo Err_PROCEDURE

10        If SingleLogFile Then
              '--- save log file to a single file
20            NazwaPlikLogu = PrefixLogFile & GetComputerName & "_" & Year(VBA.Now) & SufixLogFile
30        Else
              '--- save log file every day
40            NazwaPlikLogu = PrefixLogFile & GetComputerName & "_" & Year(VBA.Now) & Format(Month(VBA.Now), "00") & Format(Day(VBA.Now), "00") & SufixLogFile
50        End If

Exit_PROCEDURE:
    Exit Function

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

End Function


Public Function SciezkaDoPlikuLogaLvw()

    Dim strNazwa As String
    Dim strPath As String

      Const sfName = "SciezkaDoPlikuLogaLvw"
      Dim ErrNumber, ErrDescription
On Error GoTo Err_PROCEDURE

10        strPath = Application.CurrentProject.Path
20        strNazwa = PrefixLogFile & GetComputerName & "_" & Year(VBA.Now) & "_lvw" & SufixLogFile

30        SciezkaDoPlikuLogaLvw = strPath & "\" & strNazwa

Exit_PROCEDURE:
    Exit Function

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

End Function
