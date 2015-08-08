VERSION 5.00
Begin VB.Form frmEditFileEntry 
   Caption         =   "Edit File Entry"
   ClientHeight    =   7335
   ClientLeft      =   765
   ClientTop       =   870
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   10275
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnOpenPDF 
      Caption         =   "&Open PDF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton btnLast 
      Caption         =   "&Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton btnPrevious 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton btnFirst 
      Caption         =   "&First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Details:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   1440
      TabIndex        =   18
      Top             =   840
      Width           =   9135
      Begin VB.TextBox txtNotes 
         Height          =   1215
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox txtFilepath 
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1080
         Width           =   6975
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
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
         Left            =   3120
         TabIndex        =   11
         Top             =   5040
         Width           =   1095
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
         Left            =   4440
         TabIndex        =   12
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox txtArticleTitle 
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1560
         Width           =   6975
      End
      Begin VB.TextBox txtJournalTitle 
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   4
         Top             =   2040
         Width           =   6975
      End
      Begin VB.TextBox txtJournalSubject 
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2520
         Width           =   6975
      End
      Begin VB.TextBox txtFirstAuthor 
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   6
         Top             =   3000
         Width           =   6975
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   7
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtVolNum 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   8
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtPageNum 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtFilename 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   1
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Filepath:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Article Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Journal Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Journal Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "First Author:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Volumn No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Page No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label lblFileStatus 
      Caption         =   "File availability is shown here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Label lblRecSummary 
      Alignment       =   1  'Right Justify
      Caption         =   "Record Summary displayed here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   28
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Edit File Entry"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmEditFileEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################################################
Private m_objPDFDetails As PDFFileDetails
Private m_searchResults As Collection
Private m_curr_index As Integer
'#################################################################################
Private Declare Function ShellExecute _
  Lib "shell32.dll" _
  Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) _
  As Long
'#################################################################################
Private Sub btnOpenPDF_Click()
  Dim strFile As String
  Dim strAction As String
  Dim lngErr As Long

  ' Edit this:
  Dim sFilePath As String
  sFilePath = m_objPDFDetails.m_filepath
  If (InStrRev(sFilePath, "\") = Len(sFilePath)) Then
    'there is a "\" at the end
    strFile = sFilePath & m_objPDFDetails.m_filename
  Else
    strFile = sFilePath & "\" & m_objPDFDetails.m_filename
  End If
  strAction = "OPEN"  ' action might be OPEN, NEW or other, depending on what you need to do
  lngErr = ShellExecute(0, strAction, strFile, "", "", 0)
  If (lngErr <> 0 And lngErr <> 42) Then
    MsgBox "Could not open the PDF file.", vbCritical, getApplnName
  End If
End Sub
'#################################################################################
Public Sub init_form(ByRef p_searchResults As Collection, sFilename As String)
  Set m_searchResults = p_searchResults
  m_curr_index = 0
  'get the PDFFileDetails object from the collection based on the sFilename
  Dim i As Integer
  Set m_objPDFDetails = searchPDFFileDetailsIn(p_searchResults, sFilename)
  '-------------------------------------------------------------------------
  If m_curr_index > 0 Then    'display PDF file details now
    setFileDetails m_objPDFDetails
  End If
  '-------------------------------------------------------------------------
  setButtonsStatus
  Me.lblRecSummary.Caption = "Record: " & m_curr_index & " of " & m_searchResults.Count
  '-------------------------------------------------------------------------
  'check whether PDF is present in filepath
  setFileStatus
  '-------------------------------------------------------------------------
  Me.Show
  Me.SetFocus
  '-------------------------------------------------------------------------
End Sub
'#################################################################################
'check whether PDF is present in filepath
Private Sub setFileStatus()
  Dim fs As New FileSystemObject
  If (fs.FileExists(m_objPDFDetails.m_filepath & "\" & m_objPDFDetails.m_filename)) Then
    Me.btnOpenPDF.Enabled = True
    Me.lblFileStatus.Caption = "File present in the specified folder."
    Me.lblFileStatus.ForeColor = &HC000&     'green
  Else
    Me.btnOpenPDF.Enabled = False
    Me.lblFileStatus.Caption = "File not present in the specified folder."
    Me.lblFileStatus.ForeColor = &HFF&       'red
  End If
End Sub
'#################################################################################
Private Function searchPDFFileDetailsIn(ByRef p_searchResults As Collection, sFilename As String) As PDFFileDetails
  Dim t_pdfFileDtls As PDFFileDetails
  Set searchPDFFileDetailsIn = New PDFFileDetails
  For i = 1 To p_searchResults.Count
    Set t_pdfFileDtls = p_searchResults.Item(i)
    If sFilename = t_pdfFileDtls.m_filename Then
      m_curr_index = i
      Set searchPDFFileDetailsIn = t_pdfFileDtls
      Exit For
    End If
  Next
End Function
'#################################################################################
Private Sub setButtonsStatus()
  If m_curr_index > 0 Then
    Me.btnSave.Enabled = True
  Else
    Me.btnSave.Enabled = False
  End If
  If m_curr_index = 1 Then
    'showing first record - disable first and previous buttons
    Me.btnFirst.Enabled = False
    Me.btnPrevious.Enabled = False
  Else
    Me.btnFirst.Enabled = True
    Me.btnPrevious.Enabled = True
  End If
  If m_curr_index = m_searchResults.Count Then
    'showing last record - disable next and last buttons
    Me.btnNext.Enabled = False
    Me.btnLast.Enabled = False
  Else
    Me.btnNext.Enabled = True
    Me.btnLast.Enabled = True
  End If
End Sub
'#################################################################################
Private Sub setFileDetails(ByRef objPDFDetails As PDFFileDetails)
  Me.txtArticleTitle.text = objPDFDetails.m_article_title
  Me.txtJournalTitle.text = objPDFDetails.m_journal_title
  Me.txtJournalSubject.text = objPDFDetails.m_journal_subject
  Me.txtVolNum.text = objPDFDetails.m_volume_no
  Me.txtPageNum.text = objPDFDetails.m_page_no
  Me.txtFirstAuthor.text = objPDFDetails.m_first_author
  Me.txtYear.text = objPDFDetails.m_year
  Me.txtFilename.text = objPDFDetails.m_filename
  Me.txtFilepath.text = objPDFDetails.m_filepath
  Me.txtNotes.text = objPDFDetails.m_notes
End Sub
'#################################################################################
Private Sub btnClose_Click()
  Me.Hide
End Sub
'#################################################################################
Private Sub btnFirst_Click()
  m_curr_index = 1
  Dim t_pdfDetail As PDFFileDetails
  Set t_pdfDetail = m_searchResults.Item(m_curr_index)
  init_form m_searchResults, t_pdfDetail.m_filename
End Sub
'#################################################################################
Private Sub btnPrevious_Click()
  m_curr_index = m_curr_index - 1
  If (m_curr_index < 1) Then
    m_curr_index = 1
  End If
  Dim t_pdfDetail As PDFFileDetails
  Set t_pdfDetail = m_searchResults.Item(m_curr_index)
  init_form m_searchResults, t_pdfDetail.m_filename
End Sub
'#################################################################################
Private Sub btnNext_Click()
  m_curr_index = m_curr_index + 1
  If (m_curr_index > m_searchResults.Count) Then
    m_curr_index = m_searchResults.Count
  End If
  Dim t_pdfDetail As PDFFileDetails
  Set t_pdfDetail = m_searchResults.Item(m_curr_index)
  init_form m_searchResults, t_pdfDetail.m_filename
End Sub
'#################################################################################
Private Sub btnLast_Click()
  m_curr_index = m_searchResults.Count
  Dim t_pdfDetail As PDFFileDetails
  Set t_pdfDetail = m_searchResults.Item(m_curr_index)
  init_form m_searchResults, t_pdfDetail.m_filename
End Sub
'#################################################################################
Private Sub btnSave_Click()
  Dim sFileDetailsSave As String
  Dim sErr As String
  Dim intCount As Integer
  sFileDetailsSave = getFileSaveSQL
  intCount = DatabaseMod.executeUpdate(sErr, sFileDetailsSave)
  If (intCount <= 0 Or Len(Trim(sErr)) > 0) Then
    MsgBox "File Details could not be saved due to following error: " & vbCrLf & sErr, vbCritical, getApplnName
  Else
    m_objPDFDetails.m_article_title = Me.txtArticleTitle.text
    m_objPDFDetails.m_filepath = Me.txtFilepath.text
    m_objPDFDetails.m_first_author = Me.txtFirstAuthor.text
    m_objPDFDetails.m_journal_subject = Me.txtJournalSubject.text
    m_objPDFDetails.m_journal_title = Me.txtJournalTitle.text
    m_objPDFDetails.m_page_no = Me.txtPageNum.text
    m_objPDFDetails.m_volume_no = Me.txtVolNum.text
    m_objPDFDetails.m_year = Me.txtYear.text
    m_objPDFDetails.m_notes = Me.txtNotes.text
    MsgBox "File Details saved.", vbInformation, getApplnName
  End If
End Sub
'#################################################################################
Private Function getFileSaveSQL() As String
  Dim sSQL As String
  sSQL = "UPDATE  PDFFileDetails " & _
         "SET     PDFFileDetails.article_title = """ & Replace(Me.txtArticleTitle.text, "'", "''") & """, " & _
         "        PDFFileDetails.journal_title = """ & Replace(Me.txtJournalTitle.text, "'", "''") & """, " & _
         "        PDFFileDetails.journal_subject = """ & Replace(Me.txtJournalSubject.text, "'", "''") & """, " & _
         "        PDFFileDetails.[year] = """ & Replace(Me.txtYear.text, "'", "''") & """, " & _
         "        PDFFileDetails.volume_no = """ & Replace(Me.txtVolNum.text, "'", "''") & """, " & _
         "        PDFFileDetails.page_no = """ & Replace(Me.txtPageNum.text, "'", "''") & """, " & _
         "        PDFFileDetails.first_author = """ & Replace(Me.txtFirstAuthor.text, "'", "''") & """, " & _
         "        PDFFileDetails.filepath = """ & Replace(Me.txtFilepath.text, "'", "''") & """, " & _
         "        PDFFileDetails.notes = """ & Replace(Me.txtNotes.text, "'", "''") & """ " & _
         "WHERE   PDFFileDetails.filename = """ & m_objPDFDetails.m_filename & """"
  '------------------------------------------
  getFileSaveSQL = sSQL
  '------------------------------------------
End Function
'#################################################################################
