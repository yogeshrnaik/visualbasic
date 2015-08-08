VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   8085
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   11295
      Begin VB.TextBox txtFilepath 
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtNotes 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Case Sensitive"
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
         Left            =   8400
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Logical Operation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8400
         TabIndex        =   25
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton radLogicalAnd 
            Caption         =   "&AND"
            Height          =   195
            Left            =   480
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton radLogicalOr 
            Caption         =   "&OR"
            Height          =   195
            Left            =   480
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   6000
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "C&lear"
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
         Left            =   5040
         TabIndex        =   15
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtPageNum 
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtVolNum 
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   6000
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFirstAuthor 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtJournalSubject 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtJournalTitle 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtArticleTitle 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "&Search"
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
         Left            =   3600
         TabIndex        =   14
         Top             =   3120
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
         Left            =   6480
         TabIndex        =   16
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Left            =   4800
         TabIndex        =   30
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         Left            =   1440
         TabIndex        =   29
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Search"
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
         Height          =   255
         Left            =   4740
         TabIndex        =   27
         Top             =   240
         Width           =   1695
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
         Left            =   4800
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
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
         Left            =   4800
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
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
         Left            =   4680
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
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
         Left            =   4920
         TabIndex        =   22
         Top             =   720
         Width           =   975
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
         Left            =   960
         TabIndex        =   21
         Top             =   2160
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
         Left            =   720
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
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
         Left            =   960
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
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
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvSearchResults 
      Height          =   3375
      Left            =   360
      TabIndex        =   17
      Top             =   4320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Article Title"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Journal Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Journal Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "First Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Volumn No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Page No."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblSearchResult 
      Caption         =   "No. of files found:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   2535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################################################
Private m_searchResults As Collection
'#################################################################################
Private Sub btnCancel_Click()
  Me.Hide
End Sub
'#################################################################################
Private Sub btnClear_Click()
  Me.txtArticleTitle.text = ""
  Me.txtFilename.text = ""
  Me.txtFirstAuthor.text = ""
  Me.txtJournalSubject.text = ""
  Me.txtJournalTitle.text = ""
  Me.txtPageNum.text = ""
  Me.txtVolNum.text = ""
  Me.txtYear.text = ""
  Me.txtNotes.text = ""
  Me.txtFilepath.text = ""
  Me.txtArticleTitle.SetFocus
End Sub
'#################################################################################
Private Sub btnClose_Click()
  Me.Hide
End Sub
'#################################################################################
Private Sub find()
  Dim sSearchSQL As String
  Dim intResultCount As Integer
  Dim rs As ADODB.Recordset
  sSearchSQL = getSearchQuery
  Set rs = DatabaseMod.executeQuery(sSearchSQL)
  Set m_searchResults = New Collection
  Me.lvSearchResults.ListItems.Clear
  intResultCount = rs.RecordCount
  If (intResultCount = 0) Then
    lblSearchResult.ForeColor = &HFF&
  Else
    lblSearchResult.ForeColor = &HC00000
  End If
  lblSearchResult.Caption = "No. of files found: " & intResultCount
  
  If (intResultCount = 0) Then
    Exit Sub
  End If
  
  rs.MoveFirst
  Dim lstItem As ListItem
  While Not rs.EOF
    With Me.lvSearchResults
      Dim pdfFileDtls As PDFFileDetails
      Set pdfFileDtls = New PDFFileDetails
      pdfFileDtls.init rs("filename"), rs("filepath"), rs("article_title"), _
                        rs("journal_title"), rs("journal_subject"), rs("year"), _
                        rs("volume_no"), rs("page_no"), rs("first_author"), _
                        IIf(IsNull(rs("notes")), "", rs("notes"))
      Set lstItem = .ListItems.Add(, , rs("filename"))
      lstItem.SubItems(1) = rs("article_title")
      lstItem.SubItems(2) = rs("journal_title")
      lstItem.SubItems(3) = rs("journal_subject")
      lstItem.SubItems(4) = rs("first_author")
      lstItem.SubItems(5) = rs("year")
      lstItem.SubItems(6) = rs("volume_no")
      lstItem.SubItems(7) = rs("page_no")
      m_searchResults.Add pdfFileDtls, pdfFileDtls.m_filename
    End With
    rs.MoveNext
  Wend
  DatabaseMod.closeRecordSet rs
End Sub
'#################################################################################
Private Sub btnSearch_Click()
  find
End Sub
'#################################################################################
Private Function getSearchQuery() As String
  Dim sSQL As String
  sSQL = "SELECT PDFFileDetails.article_title, " & _
         "PDFFileDetails.journal_title, " & _
         "PDFFileDetails.journal_subject, " & _
         "PDFFileDetails.[year], " & _
         "PDFFileDetails.volume_no, " & _
         "PDFFileDetails.page_no, " & _
         "PDFFileDetails.first_author, " & _
         "PDFFileDetails.filename, " & _
         "PDFFileDetails.filepath, " & _
         "PDFFileDetails.notes " & _
         "FROM PDFFileDetails "
  '------------------------------------------
  Dim sLogicalOp As String
  If (Me.radLogicalAnd.Value = True) Then
    sSQL = sSQL & "WHERE 1 = 1 "
    sLogicalOp = " AND "
  Else
    sSQL = sSQL & "WHERE 1 = 2 "
    sLogicalOp = " OR "
  End If
  '------------------------------------------
'  Dim blnCaseSensitive As Boolean
'  blnCaseSensitive = False
'  If (chkCase.Value = 1) Then
'    blnCaseSensitive = True
'  End If
  '------------------------------------------
  'Search on Article Title
  If (Len(Trim(Me.txtArticleTitle.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.article_title) like ""%" & _
                  UCase$((Me.txtArticleTitle.text)) & "%"""
  End If
  '------------------------------------------
  'Search on Filename
  If (Len(Trim(Me.txtFilename.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.filename like ""%" & _
                  Trim(Me.txtFilename.text) & "%"""
  End If
  '------------------------------------------
  'Search on First Author
  If (Len(Trim(Me.txtFirstAuthor.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.first_author like ""%" & _
                  Trim(Me.txtFirstAuthor.text) & "%"""
  End If
  '------------------------------------------
  'Search on Journal Subject
  If (Len(Trim(Me.txtJournalSubject.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.journal_subject like ""%" & _
                  Trim(Me.txtJournalSubject.text) & "%"""
  End If
  '------------------------------------------
  'Search on Journal Title
  If (Len(Trim(Me.txtJournalTitle.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.journal_title like ""%" & _
                  Trim(Me.txtJournalTitle.text) & "%"""
  End If
  '------------------------------------------
  'Search on Page Number
  If (Len(Trim(Me.txtPageNum.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.page_no like ""%" & _
                  Trim(Me.txtPageNum.text) & "%"""
  End If
  '------------------------------------------
  'Search on Volume Number
  If (Len(Trim(Me.txtVolNum.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.volume_no like ""%" & _
                  Trim(Me.txtVolNum.text) & "%"""
  End If
  '------------------------------------------
  'Search on Year
  If (Len(Trim(Me.txtYear.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.[year] like ""%" & _
                  Trim(Me.txtYear.text) & "%"""
  End If
  '------------------------------------------
  'Search on Notes
  If (Len(Trim(Me.txtNotes.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.notes) like ""%" & _
                  UCase$((Me.txtNotes.text)) & "%"""
  End If
  '------------------------------------------
  'Search on File path
  If (Len(Trim(Me.txtFilepath.text)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.filepath) like ""%" & _
                  UCase$((Me.txtFilepath.text)) & "%"""
  End If
  '------------------------------------------
  sSQL = sSQL & " ORDER BY PDFFileDetails.article_title ASC;"
  getSearchQuery = sSQL
End Function
'#################################################################################
Private Sub Form_Load()
  find
End Sub
'#################################################################################
Private Sub lvSearchResults_DblClick()
  Dim sFilename As String
  If (Me.lvSearchResults.ListItems.Count > 0) Then
    sFilename = Me.lvSearchResults.SelectedItem.text
    frmEditFileEntry.init_form m_searchResults, sFilename
  End If
End Sub
'#################################################################################
Private Sub radLogicalAnd_Click()
  find
End Sub
'#################################################################################
Private Sub radLogicalOr_Click()
  find
End Sub
'#################################################################################
Private Sub txtArticleTitle_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtFilename_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtFilepath_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtFirstAuthor_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtJournalTitle_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtJournalSubject_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtNotes_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtPageNum_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub txtVolNum_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################
Private Sub chkCase_Click()
  find
End Sub
'#################################################################################
Private Sub txtYear_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Then
    find
  End If
End Sub
'#################################################################################

