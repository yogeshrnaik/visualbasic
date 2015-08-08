VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   8085
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   9600
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnSaveToExcel 
      Caption         =   "Save to &Excel"
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
      Left            =   10200
      TabIndex        =   33
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Select &All"
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
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   0
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
         TabIndex        =   29
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
         Height          =   1215
         Left            =   8400
         TabIndex        =   26
         Top             =   600
         Width           =   1815
         Begin VB.OptionButton radLogicalAnd 
            Caption         =   "AN&D"
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
         Top             =   3360
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
         Top             =   3360
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
         Top             =   3360
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvSearchResults 
      Height          =   3495
      Left            =   360
      TabIndex        =   18
      Top             =   4440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6165
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlColumnHeaderIcons"
      SmallIcons      =   "imlColumnHeaderIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "First Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Year"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Volumn No."
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Page No."
         Object.Width           =   2364
      EndProperty
   End
   Begin MSComctlLib.ImageList imlColumnHeaderIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
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
      Left            =   2160
      TabIndex        =   0
      Top             =   4080
      Width           =   2895
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################################################
Option Explicit
Private m_searchResults As Collection
Private TT As CTooltip
Private m_lCurItemIndex As Long
'#################################################################################
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'#################################################################################
Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18
'#################################################################################
Private Type POINTAPI
    x As Long
    y As Long
End Type
'#################################################################################
Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type
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
  Dim isAndOpr As Boolean
  isAndOpr = Me.radLogicalAnd.Value
  sSearchSQL = getSearchQuery(isAndOpr, Me.txtArticleTitle.text, Me.txtFilename.text, _
                                Me.txtFirstAuthor.text, Me.txtJournalSubject.text, _
                                Me.txtJournalTitle.text, Me.txtPageNum.text, _
                                Me.txtVolNum.text, Me.txtYear.text, _
                                Me.txtNotes.text, Me.txtFilepath.text)
  Set rs = DatabaseMod.executeQuery(sSearchSQL)
  Set m_searchResults = New Collection
  Me.lvSearchResults.ListItems.Clear
  intResultCount = rs.RecordCount
  If (intResultCount = 0) Then
    lblSearchResult.ForeColor = &HFF& 'red
    Me.chkSelectAll.Enabled = False
    Me.btnSaveToExcel.Enabled = False
  Else
    lblSearchResult.ForeColor = &HC00000 'blue
    Me.chkSelectAll.Enabled = True
    Me.btnSaveToExcel.Enabled = True
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
      Set lstItem.Tag = pdfFileDtls
    End With
    rs.MoveNext
  Wend
  'List BackColour Formatting
  'FormatList.SetListViewColor Me.lvSearchResults, Me.Picture1, vbWhite, vbLightGreen
  FormatList
  DatabaseMod.closeRecordSet rs
End Sub

'#################################################################################
Private Sub btnSearch_Click()
  find
End Sub
'#################################################################################
Private Sub chkSelectAll_Click()
  Dim i As Integer
  If Me.chkSelectAll.Value = CheckBoxConstants.vbChecked Then
    For i = 1 To Me.lvSearchResults.ListItems.Count
      Me.lvSearchResults.ListItems.Item(i).Checked = True
    Next
  Else
    For i = 1 To Me.lvSearchResults.ListItems.Count
      Me.lvSearchResults.ListItems.Item(i).Checked = False
    Next
  End If
End Sub
'#################################################################################
Private Sub Form_GotFocus()
  refreshList
End Sub
'#################################################################################
Private Sub refreshList()
  Dim i As Integer
  For i = 1 To Me.lvSearchResults.ListItems.Count
    Dim t_pdf As PDFFileDetails
    Dim lstItem As ListItem
    Set lstItem = Me.lvSearchResults.ListItems.Item(i)
    Set t_pdf = lstItem.Tag
    lstItem.text = t_pdf.m_filename
    lstItem.SubItems(1) = t_pdf.m_article_title
    lstItem.SubItems(2) = t_pdf.m_journal_title
    lstItem.SubItems(3) = t_pdf.m_journal_subject
    lstItem.SubItems(4) = t_pdf.m_first_author
    lstItem.SubItems(5) = t_pdf.m_year
    lstItem.SubItems(6) = t_pdf.m_volume_no
    lstItem.SubItems(7) = t_pdf.m_page_no
  Next
  FormatList
End Sub
'#################################################################################

Private Sub Form_Load()
  find
  ListHeaders
  Me.lvSearchResults.ColumnHeaders.Item(2).Icon = ASCENDING
  Set TT = New CTooltip
  TT.Style = TTBalloon
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
Private Sub lvSearchResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = MouseButtonConstants.vbRightButton Then
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim lstItem As ListItem
    Dim t_pdf  As PDFFileDetails
    
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(Me.lvSearchResults.hwnd, LVM_HITTEST, 0, lvhti) + 1
    If (lItemIndex <> 0) Then
      Me.lvSearchResults.SelectedItem = Me.lvSearchResults.ListItems.Item(lItemIndex)
      Dim i As Integer
      Dim isChecked As Boolean
      isChecked = False
      For i = 1 To Me.lvSearchResults.ListItems.Count
        If (Me.lvSearchResults.ListItems.Item(i).Checked) Then
          isChecked = True
          Exit For
        End If
      Next
      If (isChecked = True) Then
        'atleast one is checked, hence enable all menu items
        mdiMain.mMoveFile.Enabled = True
        mdiMain.mMoveChecked.Enabled = True
        mdiMain.mDeleteFile.Enabled = True
        mdiMain.mDeleteChecked.Enabled = True
      Else
        mdiMain.mMoveFile.Enabled = True
        mdiMain.mMoveChecked.Enabled = False
        mdiMain.mDeleteFile.Enabled = True
        mdiMain.mDeleteChecked.Enabled = False
      End If
      Me.PopupMenu mdiMain.mContextMenu
    End If
  End If
End Sub
'#################################################################################
Private Sub lvSearchResults_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lvhti As LVHITTESTINFO
  Dim lItemIndex As Long
  Dim lstItem As ListItem
  Dim t_pdf  As PDFFileDetails
  
  lvhti.pt.x = x / Screen.TwipsPerPixelX
  lvhti.pt.y = y / Screen.TwipsPerPixelY
  lItemIndex = SendMessage(Me.lvSearchResults.hwnd, LVM_HITTEST, 0, lvhti) + 1
  
  If m_lCurItemIndex <> lItemIndex Then
    m_lCurItemIndex = lItemIndex
    If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
      TT.Destroy
    Else
      Set lstItem = Me.lvSearchResults.ListItems(m_lCurItemIndex)
      Set t_pdf = lstItem.Tag
      TT.Title = t_pdf.m_filename
      If (GeneralMod.doesFileExist(t_pdf.m_filepath, t_pdf.m_filename)) Then
        TT.Icon = TTIconInfo
        TT.ForeColor = &HC00000
        TT.TipText = "File present at" & vbNewLine & "'" & t_pdf.m_filepath & "'."
        TT.Create Me.lvSearchResults.hwnd
      Else
        TT.Icon = TTIconError
        TT.ForeColor = vbRed
        TT.TipText = "File not found at" & vbNewLine & "'" & t_pdf.m_filepath & "'."
        TT.Create Me.lvSearchResults.hwnd
      End If
    End If
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
Private Sub lvSearchResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  '------------------------------------------------------------------
  Dim colVar As ColumnHeader
  ' If the ListView is already sorted by the clicked column, _
  ' just reverse the order. Otherwise, sort the clicked column ascending.
  If Me.lvSearchResults.Sorted = True And ColumnHeader.SubItemIndex = Me.lvSearchResults.SortKey Then
    If Me.lvSearchResults.SortOrder = lvwAscending Then
      Me.lvSearchResults.SortOrder = lvwDescending
    Else
      Me.lvSearchResults.SortOrder = lvwAscending
    End If
  Else
    Me.lvSearchResults.Sorted = True
    Me.lvSearchResults.SortKey = ColumnHeader.SubItemIndex
    Me.lvSearchResults.SortOrder = lvwAscending
  End If
  '------------------------------------------------------------------
  'Now, use the sort information to update the up or down arrows on the columnheader
  For Each colVar In Me.lvSearchResults.ColumnHeaders
    If colVar.SubItemIndex = Me.lvSearchResults.SortKey Then
      If Me.lvSearchResults.SortOrder = lvwDescending Then
        colVar.Icon = DESCENDING
      Else
        colVar.Icon = ASCENDING
      End If
    Else
      colVar.Icon = NO_ORDER
    End If
  Next colVar
  '------------------------------------------------------------------
  're-init the m_searchResults after sorting
'  'firstly remove all PDFFileDetails from search results collection
'  While m_searchResults.Count >= 1
'    m_searchResults.Remove (1)
'  Wend
  Set m_searchResults = New Collection
  '------------------------------------------------------------------
  'add the sorted PDFFileDetails to the search results collection
  Dim i As Integer
  i = 1
  For i = 1 To Me.lvSearchResults.ListItems.Count
    Dim t_pdf As PDFFileDetails
    Set t_pdf = Me.lvSearchResults.ListItems.Item(i).Tag
    'Me.txtArticleTitle.text = Me.txtArticleTitle.text & "," & t_pdf.m_filename
    m_searchResults.Add t_pdf
  Next
  '------------------------------------------------------------------
End Sub
'#################################################################################
Private Sub ListHeaders()
  Dim imlItem As ListImage
 'Set up ImageList for ColumnHeaders
  imlColumnHeaderIcons.ImageHeight = 12
  imlColumnHeaderIcons.ImageWidth = 12
  Set imlItem = imlColumnHeaderIcons.ListImages.Add(, , LoadResPicture(101, vbResBitmap))
  Set imlItem = imlColumnHeaderIcons.ListImages.Add(, , LoadResPicture(102, vbResBitmap))
  'Set up ListView
  Me.lvSearchResults.View = lvwReport
  Me.lvSearchResults.ColumnHeaderIcons = imlColumnHeaderIcons
  Me.lvSearchResults.Arrange = lvwAutoTop
  Me.lvSearchResults.LabelEdit = lvwManual
End Sub
'#################################################################################
Private Sub FormatList()
  Dim lstItem As ListItem
  Dim Counter As Long
  Dim t_pdf As PDFFileDetails
 ' Set the variable to the ListItem.
  For Counter = 1 To Me.lvSearchResults.ListItems.Count
    Set lstItem = Me.lvSearchResults.ListItems.Item(Counter)
    Set t_pdf = lstItem.Tag
    If (GeneralMod.doesFileExist(t_pdf.m_filepath, t_pdf.m_filename)) Then
      'Me.lvSearchResults.ListItems.item(Counter).ListSubItems(3).ForeColor = vbRed
      formatListItem lstItem, &HC00000, False
    Else
      formatListItem lstItem, vbRed, True
    End If
  Next Counter
End Sub
'#################################################################################
Private Sub formatListItem(ByRef lstItem As ListItem, ForeColor As Long, isBold As Boolean)
  lstItem.ForeColor = ForeColor
  lstItem.Bold = isBold
  Dim i As Integer
  For i = 1 To lstItem.ListSubItems.Count
    lstItem.ListSubItems.Item(i).ForeColor = ForeColor
    lstItem.ListSubItems.Item(i).Bold = isBold
  Next
End Sub
'#################################################################################
Private Sub btnSaveToExcel_Click()
On Error GoTo Hell
  Dim xlApp As Excel.Application
  Dim xlBook As Excel.Workbook
  Dim xlSheet As Excel.Worksheet
  Dim oRng As Excel.Range
  '--------------------------------------------------------------------
  'get the excel file name from user
  Dim xlsFilename As String
  Me.FileDialog.Filter = "MS Excel Files (*.xls)|*.xls;"
  Me.FileDialog.CancelError = True
  Me.FileDialog.DialogTitle = "Save File as"
  Me.FileDialog.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY

  Me.FileDialog.ShowSave
  xlsFilename = Me.FileDialog.FileName
  If (Len(xlsFilename) = 0) Then
    Exit Sub
  Else
    'delete any previous file with same name
    Dim fs As New FileSystemObject
    If (fs.FileExists(xlsFilename)) Then
      fs.DeleteFile xlsFilename, True
    End If
    Set fs = Nothing
  End If
  '--------------------------------------------------------------------
  Set xlApp = CreateObject("Excel.Application")
  Set xlBook = xlApp.Workbooks.Add
  Set xlSheet = xlBook.Worksheets(1)
  '--------------------------------------------------------------------
  'formatting excel sheet
  Set oRng = xlSheet.Range("C1", "C1") 'set width of column containing Ariticle Title
  oRng.EntireColumn.ColumnWidth = 50
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("D1", "D1") 'set width of column containing Journal Title
  oRng.EntireColumn.ColumnWidth = 30
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("E1", "E1") 'set width of column containing Journal Subject
  oRng.EntireColumn.ColumnWidth = 15
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("F1", "F1") 'set width of column containing First Author
  oRng.EntireColumn.ColumnWidth = 15
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("G1", "G1") 'set width of column containing Year
  oRng.EntireColumn.ColumnWidth = 10
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("H1", "H1") 'set width of column containing Volume No.
  oRng.EntireColumn.ColumnWidth = 12
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  Set oRng = xlSheet.Range("I1", "I1") 'set width of column containing Page No.
  oRng.EntireColumn.ColumnWidth = 10
  oRng.EntireColumn.WrapText = True
  '--------------------------------------------------------------------
  'formatting the excel sheet
  With xlSheet.Application
    ' The following statement shows the sheet.
    .Visible = True
    'add column headers in excel
    .Cells(1, 1) = "Sr. No."
    .Cells(1, 2) = "Filename"
    .Cells(1, 3) = "Article Title"
    .Cells(1, 4) = "Journal Title"
    .Cells(1, 5) = "Journal Subject"
    .Cells(1, 6) = "First Author"
    .Cells(1, 7) = "Year"
    .Cells(1, 8) = "Volume No."
    .Cells(1, 9) = "Page No."
    '--------------------------------------------------------------------
    'alignment for Sr. No.
    Set oRng = xlSheet.Range("A1", "A1")
    With oRng
      .Font.Bold = True
      .EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
      .EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop 'verticle alignement as Top
    End With
    '--------------------------------------------------------------------
    'Format columns from Filename to First Author (B1:F1) as bold, horizontal alignment as left
    With xlSheet.Range("B1", "F1")
      .Font.Bold = True
      .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
      .EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop 'verticle alignement as Top
    End With
    '--------------------------------------------------------------------
    'for Year, Vol. No. and Page No.
    With xlSheet.Range("G1", "I1")
      
      .Font.Bold = True
      .EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
      .EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop 'verticle alignement as Top
    End With
    '--------------------------------------------------------------------
    'put search results on the worksheet
    Dim i As Integer
    For i = 1 To Me.lvSearchResults.ListItems.Count
      Dim t_pdf As PDFFileDetails
      Set t_pdf = Me.lvSearchResults.ListItems.Item(i).Tag
      .Cells(i + 1, 1) = i
      .Cells(i + 1, 2) = t_pdf.m_filename
      '.Cells(i + 1, 2).AddComment (t_pdf.m_filepath)
      .Cells(i + 1, 3) = t_pdf.m_article_title
      .Cells(i + 1, 4) = t_pdf.m_journal_title
      .Cells(i + 1, 5) = t_pdf.m_journal_subject
      .Cells(i + 1, 6) = t_pdf.m_first_author
      .Cells(i + 1, 7) = "'" & t_pdf.m_year
      .Cells(i + 1, 8) = "'" & t_pdf.m_volume_no
      .Cells(i + 1, 9) = "'" & t_pdf.m_page_no
    Next i
    '--------------------------------------------------------------------
    'AutoFit column containing file name
    Set oRng = xlSheet.Range("B1", "B1")
    oRng.EntireColumn.AutoFit
    '--------------------------------------------------------------------
    'apply cell borders
    Set oRng = xlSheet.Range("A1", "I" & Me.lvSearchResults.ListItems.Count + 1)
    oRng.Borders.Value = True
    '--------------------------------------------------------------------
    ' save the workbook
    xlBook.SaveAs (xlsFilename)
    ' close the workbook.
    '.Quit
  End With
  Exit Sub
Hell:
  'when user clicks on Cancel button of File Dialog box,
  'error no. 32755 is generated. Ignore this error.
  If (Err.Number <> 32755) Then
    If (Err.Number = 70) Then 'permission denied
      'most like the file is already open
      MsgBox Err.Description & vbCrLf & _
             "Most likely the file '" & xlsFilename & "' is already open." & vbCrLf & _
             "Close the file and then try again.", vbCritical, GeneralMod.getApplnName
    Else
      MsgBox "Error occured while saving to Excel:" & vbCrLf & _
             "Error number: " & Err.Number & vbCrLf & _
             "Error Description: " & Err.Description, vbCritical, GeneralMod.getApplnName
    End If
  End If
  Set oRng = Nothing
  Set xlSheet = Nothing
  Set xlBook = Nothing
  Set xlApp = Nothing
End Sub
'#################################################################################

