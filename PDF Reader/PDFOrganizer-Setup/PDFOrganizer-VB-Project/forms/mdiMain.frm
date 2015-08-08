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
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mAbout_Click()
  Dim sAbout As String
  sAbout = "PDF Organizer 1.0" & vbCr
  sAbout = sAbout & "Developed By: Yogesh R Naik" & vbCr
  sAbout = sAbout & "Purpose: Maintains the data for PDF files." & vbCr
  sAbout = sAbout & vbTab & "Search facility to find desired article based on various " & vbCr
  sAbout = sAbout & vbTab & "user defined search criteria." & vbCr
  MsgBox sAbout, vbOKOnly, GeneralMod.getApplnName
End Sub

Private Sub mContents_Click()
  frmHelp.init_form
  frmHelp.SetFocus
End Sub

Private Sub MDIForm_Load()
  Dim db As New FileSystemObject
  If Not db.FileExists(App.Path & "\database\pdfFilesInfo.mdb") Then
    MsgBox "Database file pdfFilesInfo.mdb not found." & vbNewLine & _
           "Please make sure that pdfFilesInfo.mdb file is present in: " & _
           App.Path & "\database" & " folder.", vbCritical, GeneralMod.getApplnName
    End 'exit application
  End If
  'frmCaptureText.Show
  frmLoadPDF.Show
  'frmEditFileEntry.Show
  
End Sub

Private Sub MDIForm_Resize()
  'Me.WindowState = 2
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


Private Sub mSearch_Click()
  frmSearch.Show
  frmSearch.SetFocus
End Sub
