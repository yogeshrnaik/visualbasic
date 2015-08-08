VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCaptureText 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Extract Text"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00AC7A3E&
      BorderStyle     =   0  'None
      ForeColor       =   &H00AC7A3E&
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   11655
      Begin VB.TextBox txtPageContents 
         Height          =   5895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   360
         Width           =   11415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Output Text"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   90
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10560
      TabIndex        =   5
      Top             =   600
      Width           =   1155
      Begin VB.CommandButton btnLoad 
         Caption         =   "   Load   "
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnBrowse 
      BackColor       =   &H00C0C0C0&
      Caption         =   "..."
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
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtPDFFile 
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   6855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PDF Organizer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC7A3E&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select a PDF File or Folder:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC7A3E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmCaptureText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#####################################################################################
Private Sub Form_Resize()
  frmCaptureText.WindowState = 2
End Sub
'#####################################################################################
Private Sub btnLoad_Click()
  Dim sError As String
  If Trim(Me.txtPDFFile.text) = "" Then
    MsgBox "Please select a PDF file.", vbExclamation, GeneralMod.getApplnName()
    Exit Sub
  End If
  Dim file As New FileSystemObject
  If file.FileExists(Me.txtPDFFile.text) Then
    If Not load_file(sError, Me.txtPDFFile.text) Then
      MsgBox sError, vbCritical, getApplnName
    End If
  ElseIf file.FolderExists(Me.txtPDFFile.text) Then
    If Not load_folder(sError, Me.txtPDFFile.text) Then
      MsgBox sError, vbCritical, getApplnName
    End If
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
  Dim iStatus As Integer
  Dim bIsSuccess As Boolean
  Dim oPDFDetails As New PDFFileDetails
  '---------------------------------------------------------------------
  iStatus = PDFParserMod.extractWords(sp_error, Me.txtPDFFile.text)
  If iStatus = GeneralMod.OKAY Then
    'PDFParserMod.displayExtractedWords
    Set oPDFDetails = PDFParserMod.parsePDFContents(sp_error, Me.txtPDFFile.text)
    If sp_error <> "" Then
      bIsSuccess = False
    ElseIf Not oPDFDetails.save_in_database(sp_error) Then
      bIsSuccess = False
    Else
      bIsSuccess = True
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
  End If
  load_file = bIsSuccess
  '---------------------------------------------------------------------
End Function
'#####################################################################################
'go through specified folder and add all files details to the database
Private Function load_folder(ByRef sp_error As String, sp_folder_path As String) As Boolean
  Dim oFileSystem As New FileSystemObject
  Dim oFolder As Folder
  Dim oCurrentFile As file
  Dim oFileColl As Files
  Dim sFilePath As String
  '--------------------------------------------------------------------------
  'load all files present in the folder
  Set oFolder = oFileSystem.GetFolder(sp_folder_path)
  Set oFileColl = oFolder.Files
  If oFileColl.Count > 0 Then
    'iterate through each file present in the folder
    For Each oCurrentFile In oFileColl
      If InStr(1, StrReverse(sp_folder_path), "\", vbTextCompare) <> 0 Then
        sFilePath = sp_folder_path & oCurrentFile.Name
      Else
        sFilePath = sp_folder_path & "\" & oCurrentFile.Name
      End If
      If oFileSystem.FileExists(sFilePath) Then
        If Not load_file(sFilePath) Then
          MsgBox "Could not save details of file '" & sp_file_path & "' in database.", _
                 vbExclamation, GeneralMod.getApplnName
        End If
      End If
    Next
  End If
  '--------------------------------------------------------------------------
  Set oFileSystem = Nothing
  Set oFolder = Nothing
  Set oFileColl = Nothing
  Set oCurrentFile = Nothing
End Function
'#####################################################################################
Private Sub btnBrowse_Click()
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

