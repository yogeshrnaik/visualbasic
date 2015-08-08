VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadPDF 
   Caption         =   "Load PDF(s)"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnClearLog 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10800
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00E0E0E0&
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2520
      Width           =   11535
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton btnBrowseFolder 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10560
         Picture         =   "frmLoadPDF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Browse for folder"
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7320
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "C&ancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6120
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton btnBrowseFile 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Picture         =   "frmLoadPDF.frx":02E2
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Browse for PDF file"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPDFFile 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   960
         Width           =   7335
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "&Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4920
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkSubFolders 
         Caption         =   "Include Sub &folders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select a PDF File or Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Load PDF(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   11280
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select PDF"
   End
   Begin VB.Label lblStatus 
      Caption         =   "progress info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9855
   End
End
Attribute VB_Name = "frmLoadPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#####################################################################################
'Global variables
Dim sSeparator As String
Dim isLoading As Boolean 'flag to indicate whether the file / folder loading is in progress
Dim stopLoad As Boolean 'flag to stop the load process when user clicks on Cancel
Dim includeSubFolders As Boolean 'whether to include sub-folders
'#####################################################################################

Private Sub btnBrowseFolder_Click()
  Dim sPath As String
  sPath = BrowseFolderMod.displayFolderDialog(Me)
  If Len(sPath) > 0 Then
    Me.txtPDFFile.text = sPath
  End If
End Sub
'#####################################################################################
Private Sub btnCancel_Click()
  Me.btnCancel.Enabled = False
  notify True, "Canceling load process..."
  stopLoad = True
End Sub
'#####################################################################################
Private Sub btnClearLog_Click()
  Me.txtLog.text = ""
  Me.lblStatus.Caption = ""
End Sub
'#####################################################################################
Private Sub btnClose_Click()
  Me.Hide
End Sub
'#####################################################################################
Private Sub Form_Load()
  'Me.WindowState = 2
End Sub
'#####################################################################################
Private Sub Form_Resize()
On Error GoTo Hell
  'Me.WindowState = 2
  isLoading = False
  stopLoad = False
  Me.txtPDFFile.SetFocus
  Me.lblStatus.Caption = ""
  Me.btnCancel.Enabled = False
Hell:
  'ignore
End Sub
'#####################################################################################
'check the date of "PDFParser.dll" file
'this is required because, "PDFParser.dll" has an expiry period of 30 days from the day of file creation
Private Function checkPDFDllFileDate() As Boolean
  checkPDFDllFileDate = True
  Dim fileSys As New FileSystemObject
  Dim file As file
  Set file = fileSys.GetFile(App.path & "\bin\PDFParser.dll")
  Dim today As Date
  today = Now
  If (DateDiff("d", file.DateLastModified, today) > 20) Then
    Dim dtEndDate As Date
    dtEndDate = DateAdd("d", 20, file.DateLastModified)
    MsgBox "Please set the system date to any date between: " & vbCrLf & FormatDateTime(file.DateLastModified, vbLongDate) & _
           " and " & FormatDateTime(dtEndDate, vbLongDate), vbExclamation, getApplnName
    checkPDFDllFileDate = False
  End If
End Function
'#####################################################################################
Private Sub btnLoad_Click()
  If Not checkPDFDllFileDate Then
    Exit Sub
  End If
  Dim sError As String
  Me.lblStatus.Caption = ""
  If stopLoad Then
    Exit Sub
  End If
  If Trim(Me.txtPDFFile.text) = "" Then
    Call btnBrowseFile_Click
'    MsgBox "Please select a PDF file or" & vbCrLf & _
'          "enter path of a folder containing PDF(s).", _
'          vbExclamation, GeneralMod.getApplnName()
    Exit Sub
  End If
  Dim file As New FileSystemObject
  If file.FileExists(Me.txtPDFFile.text) Then
    If (Not isPDF(Me.txtPDFFile.text)) Then
      MsgBox "The selected file is not a PDF file.", vbExclamation, GeneralMod.getApplnName
      Exit Sub
    End If
    isLoading = True
    Me.btnCancel.Enabled = True
    load_file sError, Me.txtPDFFile.text
    Me.btnCancel.Enabled = False
    stopLoad = False
    isLoading = False
    MsgBox "Process completed.", vbInformation, getApplnName
  ElseIf file.FolderExists(Me.txtPDFFile.text) Then
    isLoading = True
    Me.btnCancel.Enabled = True
    If chkSubFolders.Value = 1 Then
      includeSubFolders = True
    Else
      includeSubFolders = False
    End If
    Dim files_loaded_count As Integer
    load_folder sError, Me.txtPDFFile.text, includeSubFolders, files_loaded_count
    Me.btnCancel.Enabled = False
    If (stopLoad) Then
      notify True, "Load process cancelled."
      stopLoad = False
    Else
      MsgBox "Process completed. No of files loaded: " & files_loaded_count, _
              vbInformation, getApplnName
    End If
    isLoading = False
  Else
    MsgBox "The path '" & Me.txtPDFFile.text & "' " & _
           "is not a valid file or folder." & vbNewLine & _
           "Please select a valid file or folder.", vbCritical, GeneralMod.getApplnName
    Exit Sub
  End If
End Sub
'#####################################################################################
'load file details in database
Private Function load_file(ByRef sp_error As String, sp_file_path As String) As Boolean
'---------------------------------------------------------------------
On Error GoTo Hell
'---------------------------------------------------------------------
  Dim iStatus As Integer
  Dim bIsSuccess As Boolean
  Dim oPDFDetails As New PDFFileDetails
  '---------------------------------------------------------------------
  'logMessage ("Loading file '" & sp_file_path & "'.")
  '---------------------------------------------------------------------
  'test
'  If (InStr(1, sp_file_path, "The biofilm matrix") > 0) Then
'    MsgBox "hi"
'  End If
  '---------------------------------------------------------------------
  iStatus = PDFParserMod.extractWords(sp_error, sp_file_path)
  If iStatus = GeneralMod.OKAY Then
    'PDFParserMod.displayExtractedWords
    Set oPDFDetails = PDFParserMod.parsePDFContents(sp_error, sp_file_path)
    If sp_error <> "" Then
      bIsSuccess = False
    ElseIf Not oPDFDetails.save_in_database(sp_error) Then
      bIsSuccess = False
    Else
      bIsSuccess = True
      logMessage ("File '" & sp_file_path & "' details saved successfully.")
    End If
  Else
    bIsSuccess = False
  End If
  '---------------------------------------------------------------------
  If bIsSuccess = False Then
    Dim sMessage As String
    sMessage = "Could not load the details of file '" & sp_file_path & "'" & vbCrLf
    sMessage = sMessage & "in the database due to following error: " & vbCrLf
    sMessage = sMessage & sp_error
    sp_error = sMessage
    logMessage (sp_error)
  End If
  load_file = bIsSuccess
  Exit Function
'--------------------------------------------------------------------------
Hell:
  sp_error = "Error occured while loading file '" & sp_file_path & "': " & vbCrLf & _
             "Error Number: " & Err.Number & vbCrLf & _
             "Error Description: " & Err.Description
  load_file = False
'--------------------------------------------------------------------------
End Function
'#####################################################################################
'go through specified folder and add all files details to the database
Private Function load_folder(ByRef sp_error As String, sp_folder_path As String, _
                             bp_includeSubFolders As Boolean, _
                             ByRef ip_file_count) As Boolean
  Dim oFileSystem As New FileSystemObject
  Dim oFolder As Folder
  Dim oCurrentFile As file
  Dim oFileColl As Files
  Dim oSubFolders As Folders
  Dim oSubFolder As Folder
  Dim sFilePath As String
  '--------------------------------------------------------------------------
On Error GoTo Hell
  '--------------------------------------------------------------------------
  DoEvents
  '--------------------------------------------------------------------------
  If (stopLoad) Then
    Exit Function
  End If
  'load all files present in the folder
  Set oFolder = oFileSystem.GetFolder(sp_folder_path)
  Set oFileColl = oFolder.Files
  If oFileColl.Count > 0 Then
    'logMessage ("Started loading files from folder '" & sp_folder_path & "'.")
    'iterate through each file present in the folder
    For Each oCurrentFile In oFileColl
      'check if the file is PDF file
      If isPDF(oCurrentFile.path) Then
        If oFileSystem.FileExists(oCurrentFile.path) Then
          If Not load_file(sp_error, oCurrentFile.path) Then
            logMessage (sp_error)
          Else
            ip_file_count = ip_file_count + 1
          End If
          DoEvents
        End If
      End If
    Next
  End If
  '--------------------------------------------------------------------------
  If (bp_includeSubFolders) Then
    'explore sub folders and load files present in those
    Set oSubFolders = oFolder.SubFolders
    If oSubFolders.Count > 0 Then
      'iterate through each sub folder present in the folder
      For Each oSubFolder In oSubFolders
        If oFileSystem.FolderExists(oSubFolder.path) Then
            If Not load_folder(sp_error, oSubFolder.path, bp_includeSubFolders, ip_file_count) Then
              logMessage (sp_error)
            End If
          End If
      Next
    End If
  End If
  '--------------------------------------------------------------------------
  Set oFileSystem = Nothing
  Set oFolder = Nothing
  Set oFileColl = Nothing
  Set oCurrentFile = Nothing
  Set oSubFolders = Nothing
  Set oSubFolder = Nothing
  load_folder = True
  Exit Function
'--------------------------------------------------------------------------
Hell:
  sp_error = "Error occured while loading folder '" & sp_folder_path & "': " & vbCrLf & _
             "Error Number: " & Err.Number & vbCrLf & _
             "Error Description: " & Err.Description
  load_folder = False
  Set oFileSystem = Nothing
  Set oFolder = Nothing
  Set oFileColl = Nothing
  Set oCurrentFile = Nothing
  Set oSubFolders = Nothing
  Set oSubFolder = Nothing
'--------------------------------------------------------------------------
End Function
'#####################################################################################
Private Sub btnBrowseFile_Click()
  Me.lblStatus.Caption = ""
  Me.FileDialog.Filter = "PDF Files (*.pdf)|*.pdf;"
  On Error Resume Next
    FileDialog.FileName = Me.txtPDFFile.text
    FileDialog.ShowOpen
    txtPDFFile.text = FileDialog.FileName
    'If Trim(Me.txtPDFFile.text) <> "" Then
    '  Call btnLoad_Click
    'end if
End Sub
'#####################################################################################
Private Sub txtPDFFile_Change()
  Dim file As New FileSystemObject
  If file.FolderExists(Me.txtPDFFile.text) Then
    'enable the checkbox
    Me.chkSubFolders.Enabled = True
  Else
    'disable the Checkbox
    Me.chkSubFolders.Enabled = False
  End If
End Sub
'#####################################################################################
Private Sub notify(success As Boolean, msg As String)
  If success Then
    Me.lblStatus.ForeColor = &HFF0000 'blue
  Else
    Me.lblStatus.ForeColor = &HFF& 'red color
  End If
  Me.lblStatus.Caption = msg
End Sub
'#####################################################################################
Private Function isPDF(sp_Filepath As String) As Boolean
On Error GoTo Hell
  'extension of the file should be PDF
  Dim extn As String
  extn = Mid(sp_Filepath, Len(sp_Filepath) - Len(".pdf") + 1, Len(".pdf"))
  If (UCase(extn) = ".PDF") Then
    isPDF = True
  Else
    isPDF = False
  End If
  Exit Function
Hell:
  isPDF = False
End Function
'#####################################################################################
Public Sub logMessage(sp_error As String)
  If (Len(sSeparator) = 0) Then
    For i = 0 To 200
      sSeparator = sSeparator & "-"
    Next
  End If
  'Me.txtLog.text = Me.txtLog.text & sp_error & vbCrLf & sSeparator & vbCrLf
  Me.txtLog.text = Me.txtLog.text & sp_error & vbCrLf
  'Me.txtLog.text = sp_error & vbCrLf
End Sub
'#####################################################################################
