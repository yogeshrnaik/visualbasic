VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "PDF Organizer"
   ClientHeight    =   5715
   ClientLeft      =   765
   ClientTop       =   1875
   ClientWidth     =   6585
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mLoadPDFs 
         Caption         =   "&Load PDF(s)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mSearch 
         Caption         =   "&Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu mLineSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mContents 
         Caption         =   "&Contents"
         Shortcut        =   ^H
         Visible         =   0   'False
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About PDF Organizer"
      End
   End
   Begin VB.Menu mContextMenu 
      Caption         =   "&Context Menu"
      Begin VB.Menu mMoveFile 
         Caption         =   "&Move File"
      End
      Begin VB.Menu mMoveChecked 
         Caption         =   "Move &Checked File(s)"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mDeleteFile 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu mDeleteChecked 
         Caption         =   "De&lete Checked File(s)"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#####################################################################################
Private Sub mAbout_Click()
'  Dim sAbout As String
'  sAbout = "PDF Organizer 1.0" & vbCr
'  sAbout = sAbout & "Developed By: Yogesh R Naik" & vbCr
'  sAbout = sAbout & "Purpose: Maintains the data for PDF files." & vbCr
'  sAbout = sAbout & vbTab & "Search facility to find desired article based on various " & vbCr
'  sAbout = sAbout & vbTab & "user defined search criteria." & vbCr
'  MsgBox sAbout, vbOKOnly, GeneralMod.getApplnName
  frmStartup.showAsAboutDialog
End Sub
'#####################################################################################
Private Sub mContents_Click()
  frmHelp.init_form
  frmHelp.SetFocus
End Sub
'#####################################################################################
Private Sub MDIForm_Load()
  Dim db As New FileSystemObject
  If Not db.FileExists(App.path & "\database\pdfFilesInfo.mdb") Then
    MsgBox "Database file pdfFilesInfo.mdb not found." & vbNewLine & _
           "Please make sure that pdfFilesInfo.mdb file is present in: " & _
           App.path & "\database" & " folder.", vbCritical, GeneralMod.getApplnName
    End 'exit application
  End If
  'frmCaptureText.Show
  frmLoadPDF.Show
  'frmEditFileEntry.Show
  Me.mContextMenu.Visible = False
End Sub
'#####################################################################################
Private Sub MDIForm_Resize()
  'Me.WindowState = 2
End Sub
'#####################################################################################
Private Sub MDIForm_Unload(Cancel As Integer)
  End 'exit application
End Sub
'#####################################################################################
Private Sub mExit_Click()
  If (MsgBox("Do you really want to exit PDF Organizer?", vbYesNo, getApplnName) = vbYes) Then
    End
  End If
End Sub
'#####################################################################################
Private Sub mLoadPDFs_Click()
  'frmCaptureText.Show
  frmLoadPDF.Show
  frmLoadPDF.SetFocus
End Sub
'#####################################################################################
Private Sub mSearch_Click()
  frmSearch.Show
  frmSearch.SetFocus
End Sub
'#####################################################################################
Private Sub mMoveFile_Click()
  Dim sNewPath As String      'excluding filename
  Dim sNewFullPath As String  'including filename
  Dim sOldPath As String      'excluding filename
  Dim sOldFullPath As String  'including filename
  Dim t_pdf As PDFFileDetails
  sNewPath = BrowseFolderMod.displayFolderDialog(Me)
  If Len(sNewPath) > 0 Then
    Set t_pdf = frmSearch.lvSearchResults.SelectedItem.Tag
    sOldPath = t_pdf.m_filepath
    If InStr(1, StrReverse(sOldPath), "\") = 1 Then
      sOldFullPath = sOldPath & t_pdf.m_filename
    Else
      sOldFullPath = sOldPath & "\" & t_pdf.m_filename
    End If
    If (MsgBox("The file '" & t_pdf.m_filename & "' " & _
                "will be moved from " & vbCrLf & "'" & sOldPath & "' to" & vbCrLf & _
                "folder '" & sNewPath & "'." & vbCrLf & _
                "Are you sure?", vbYesNo Or vbQuestion, _
                GeneralMod.getApplnName) = vbYes) Then
      'update the file path of the file in the database
      Dim sError As String
      If (t_pdf.updateFilepath(sError, sNewPath)) Then
        t_pdf.m_filepath = sNewPath
        'if success then move the file from sOldPath to sNewPath
        Dim fs As New FileSystemObject
        If (fs.FileExists(sOldFullPath)) Then
          sNewFullPath = sNewPath & "\" & t_pdf.m_filename
          fs.MoveFile sOldFullPath, sNewFullPath
        End If
        MsgBox "File moved successfully.", vbInformation, GeneralMod.getApplnName
      Else
        MsgBox sError, vbCritical, GeneralMod.getApplnName
      End If
    End If
  End If
End Sub
'#####################################################################################
Private Sub mDeleteFile_Click()
  MsgBox "Delete file: " & frmSearch.lvSearchResults.SelectedItem.text
End Sub
'#####################################################################################

