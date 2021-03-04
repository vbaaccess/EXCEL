'---------------------------------------------------------------------------------------
' Class Module  : clsExportAsXLS
' Author        : paweld
' Purpose       : A class that facilitates data export to xls
'                 - TO DO - dowolnego SQL'a
'                 - TO DO - dowolnego Rst
' Licence       : Free
'
'---------------------------------------------------------------------------------------
' Changes       :
'---------------------------------------------------------------------------------------
' 2019-01-17 , paweld : Creating
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const m_wersja = "1.2"
Private Const CurrentModuleName = "clsExportAsXLS"

Private Const fExportToXlsMsgBoxPrompt_FormName = "fExportAsXLS"
'Private m_ShowSettingWIndow As Boolean ' TO DO -

Private m_CallingModuleName As String

Private m_FilePath As String
Private m_FilePrefix As String
Private m_FileSufix As String
Private m_FileExtension As String

Private m_ExportedSQL As String
Private m_ExportedRst As Recordset

Private m_RptTitle As String

Private m_SelectDirectory As Boolean

Private m_FullPathAndFileName As String
Private m_MsgBox As String

Public Event AfterExport(sMsgBoxPrompt As String, bDaneZostalyWyexportowane As Boolean)

Private Sub Class_Initialize()
'Constructor
    m_FilePath = Application.CurrentProject.Path
    m_FilePrefix = "Export_"
    m_FileSufix = GetDateTiemString
    m_FileExtension = "xls"
    
    m_RptTitle = "Export Data"
    
    Call SetFullPathAndFileName
End Sub

Private Sub Class_Terminate()
'Destructor
End Sub

Private Function GetDateTiemString() As String
    GetDateTiemString = Format(VBA.Now, "YYYYMMDD") & "_" & Format(VBA.Now, "HHMMSS")
End Function

Public Sub INIT(CallingModuleName As String, Optional FilePath As String, Optional FilePrefix As String, Optional FileSufix As String, Optional FileExtension As String)
    m_CallingModuleName = CallingModuleName
    Call SetFullPathAndFileName
End Sub

Public Function ExportShowSettingForm(Optional MsgBox As String) As Boolean
    
    If IsLoaded(fExportToXlsMsgBoxPrompt_FormName) Then
        'do noting
    Else
        
        m_FileSufix = GetDateTiemString
        
        Access.DoCmd.OpenForm fExportToXlsMsgBoxPrompt_FormName, , , , , , CStr(objHandle)
        
        Dim objForm As Form
        Set objForm = Application.Forms(fExportToXlsMsgBoxPrompt_FormName).Form
    
        ShowWindow objForm.hWnd, 8
        objForm.SetFocus
    
    End If

End Function

Public Function Export(Optional MsgBox As String) As Boolean
    Dim bWeryfikacja As Boolean
    Dim bExport As Boolean
    Dim bAnulowano As Boolean
    
    If m_SelectDirectory Then
        Dim Path$
        Path = ""
    ' TO DO -Path = BrowseFolder("Select Catalog for export:")
        If Path = "" Then
            bAnulowano = True
        Else
            FilePath = Path
        End If
        
        If bAnulowano Then Exit Function
    End If
    
    bWeryfikacja = VerificationOfSetting(MsgBox)
    
    If bWeryfikacja Then
        bWeryfikacja = DataVerification(MsgBox)
        If bWeryfikacja Then
            FileSufix = GetDateTiemString
            bExport = ExportDataToEXCEL(MsgBox)
        End If
    End If
    
    Dim sM$
    If bExport Then
        sM = "File created" & vbLf & vbLf & m_FilePath & vbLf & m_FilePrefix & m_FileSufix & "." & m_FileExtension
    Else
        sM = "Failed to create file" & vbLf & vbLf & m_FilePath & vbLf & m_FilePrefix & m_FileSufix & "." & m_FileExtension
    End If
    
    VBA.MsgBox sM, vbInformation + vbOKOnly, "Export to EXCEL"
    
    RaiseEvent AfterExport(MsgBox, bExport)
End Function
    
Private Property Get FilePath() As String
    Call SETm_FilePrefix
    FilePath = m_FilePath
End Property

Private Property Let FilePath(ByVal newFilePath As String)
    m_FilePath = newFilePath
    Call SETm_FilePrefix
    Call SetFullPathAndFileName
End Property

Public Property Get FilePrefix() As String
    FilePrefix = m_FilePrefix
End Property

Private Property Get FileSufix() As String
    FileSufix = m_FileSufix
End Property

Private Property Let FileSufix(ByVal newFileSufix As String)
    m_FileSufix = newFileSufix
    Call SetFullPathAndFileName
End Property

Private Sub SETm_FilePrefix()
    If Len(m_FilePath) > 0 Then
        If Right(m_FilePath, 1) <> "\" Then
            m_FilePath = m_FilePath & "\"
        End If
    End If
End Sub

Public Property Let FilePrefix(ByVal newFilePrefix As String)
    m_FilePrefix = newFilePrefix
    Call SetFullPathAndFileName
End Property

Public Property Get FileExtension() As String
    FileExtension = m_FileExtension
End Property

Public Property Let FileExtension(ByVal newFileExtension As String)
    m_FileExtension = newFileExtension
End Property

Public Property Get SelectDirectory() As Boolean
    SelectDirectory = m_SelectDirectory
End Property

Public Property Let SelectDirectory(ByVal bAskForDictionary As Boolean)
    m_SelectDirectory = bAskForDictionary
End Property

Public Property Get ExportedSQL() As String
    ExportedSQL = m_ExportedSQL
End Property

Public Property Let ExportedSQL(ByVal newSQL As String)
    m_ExportedSQL = newSQL
End Property
                
Public Property Get RptTitle() As String
    RptTitle = m_RptTitle
End Property

Public Property Let RptTitle(ByVal newRptTitle As String)
    m_RptTitle = newRptTitle
End Property

Private Sub SetFullPathAndFileName()
    Dim sFilePath$
    
    sFilePath = FilePath
   
    m_FullPathAndFileName = sFilePath & m_FilePrefix & m_FileSufix & "." & m_FileExtension
End Sub

Private Function VerificationOfSetting(Optional MsgBox As String) As Boolean
    'to do optional
    VerificationOfSetting = True
End Function
                  
Private Function DataVerification(Optional MsgBox As String) As Boolean
    'to do - add aditional data werfication
    Dim iCount&
    
    If Not m_ExportedRst Is Nothing Then
        Call rstSetNothing(m_ExportedRst)
    End If
    
    If m_ExportedRst Is Nothing Then
        If Len(m_ExportedSQL) > 0 Then
            Call OpenExportedSQL(m_ExportedRst, m_ExportedSQL)
        End If
    End If
    
    If m_ExportedRst Is Nothing Then
        MsgBox = MsgBox & vbLf & "No Rst!"
    Else
        iCount = m_ExportedRst.RecordCount
        If iCount = 0 Then
            MsgBox = MsgBox & vbLf & "No data to display!!"
        Else
            DataVerification = True
        End If
    End If

End Function
                  
Private Function OpenExportedSQL(Rst As Recordset, Sql As String) As Boolean
    OpenExportedSQL = OpenRst(Rst, Sql)
End Function
                  
Private Function ExportDataToEXCEL(Optional MsgBox As String) As Boolean
    Dim bExport As Boolean
    Dim bWeryfikacja As Boolean
    
    m_FileExtension = Nz(Choose(werPliku, "csv", "xls", "xlsx"), "xls")
    
    bWeryfikacja = VerificationOfSetting(MsgBox)
    
    If bWeryfikacja Then
        bWeryfikacja = DataVerification(MsgBox)
    End If
    
    If bWeryfikacja Then
        Select Case m_FileExtension
        Case "csv"
            bExport = ExportAndSaveDataToFileCSV(MsgBox)
        Case "xls"
            bExport = ExportAndSaveDataToFileXLS(MsgBox)
        Case "xlsx"
            bExport = ExportAndSaveDataToFileXLSX(MsgBox)
        End Select
    End If
    
    Dim sM$
    If bExport Then
        sM = "The file has been created " & vbLf & vbLf & m_FilePath & vbLf & m_FilePrefix & m_FileSufix & "." & m_FileExtension
    Else
        sM = "The file could not be created" & vbLf & vbLf & m_FilePath & vbLf & m_FilePrefix & m_FileSufix & "." & m_FileExtension
    End If
    
    VBA.MsgBox sM, vbInformation + vbOKOnly, "Export to EXCEL"
    
    RaiseEvent AfterExport(MsgBox, bExport)
    
    ExportDataToEXCEL = bExport
End Function


Public Function ExportAndSaveDataToFileCSV(Optional MsgBox As String) As Boolean
    Dim bExport As Boolean

    If m_ExportedRst Is Nothing Then Exit Function
    If m_ExportedRst.RecordCount = 0 Then Exit Function
    
    '--- export to csv --- --- --- --- --- --- --- --- --- ---
    Dim strRow As String
    Dim strHead As String
    Dim hFile As Long
    Dim i%
    
    hFile = FreeFile
    Open m_FullPathAndFileName For Binary As #hFile
    
    With m_ExportedRst

        Put #hFile, , RptTitle & vbCrLf
        For i = 0 To .Fields.Count - 1
            strHead = strHead & ";" & .Fields(i).Name
        Next i
        Put #hFile, , Mid(strHead, 2) & vbCrLf
    
        .MoveFirst
        While Not .EOF
        
            strRow = vbCrLf
            For i = 0 To .Fields.Count - 1
                strRow = strRow & ";" & .Fields(i)
            Next i

            Put #hFile, , Mid(strRow, 2)
            .MoveNext
        Wend

        Close #hFile
        bExport = True
    End With

    ExportAndSaveDataToFileCSV = bExport
End Function

Public Function ExportAndSaveDataToFileXLS(Optional MsgBox As String) As Boolean
    Dim bExport As Boolean
    If m_ExportedRst Is Nothing Then Exit Function
    If m_ExportedRst.RecordCount = 0 Then Exit Function
            
    '--- export do xls --- --- --- --- --- --- --- --- --- ---
    If m_FileExtension = "xls" Then
         Dim i%
         Dim oApp As Object, owB As Object
                              '--- step 1 ---
         Set oApp = CreateObject("Excel.Application")
         oApp.Visible = False
        
                              '--- step 2 ---
         Set owB = oApp.Workbooks.Add
                              '--- step 3 ---
         owB.Sheets(1).Cells(1, 1).Value = RptTitle
         owB.Sheets(1).Cells(1, 1).Font.Bold = True
         
                              '--- step 4 ---
         For i = 0 To m_ExportedRst.Fields.Count - 1
             owB.Sheets(1).Cells(3, i + 1).Value = m_ExportedRst.Fields(i).Name
         Next
        
         owB.Sheets(1).Range("3:3").Font.Bold = True
         owB.Sheets(1).Cells(4, 1).CopyFromRecordset m_ExportedRst

                              '--- step 5 ---
         m_ExportedRst.Close
         Set m_ExportedRst = Nothing

                              '--- step 6 ---
         owB.Close SaveChanges:=True, FileName:=m_FullPathAndFileName
         Set owB = Nothing
         oApp.Quit
         Set oApp = Nothing
         bExport = True
    End If
    bExport = True
    '--- ------------- --- --- --- --- --- --- --- --- --- ---
    
    ExportAndSaveDataToFileXLS = bExport
End Function

                                                                                    
Public Function ExportAndSaveDataToFileXLSX(Optional MsgBox As String) As Boolean
    Dim bExport As Boolean
    
    If m_ExportedRst Is Nothing Then Exit Function
    If m_ExportedRst.RecordCount = 0 Then Exit Function
    
    '--- export do xlsx -- --- --- --- --- --- --- --- --- ---
    'TO DO
    '--- ------------- --- --- --- --- --- --- --- --- --- ---
    
    ExportAndSaveDataToFileXLSX = bExport
End Function
                                                                    
