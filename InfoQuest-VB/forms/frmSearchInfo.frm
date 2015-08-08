VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchInfo 
   Caption         =   "Search Information"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   1875
   ClientWidth     =   11370
   Icon            =   "frmSearchInfo.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnRefresh 
      Caption         =   "&Refresh"
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
      Left            =   7560
      TabIndex        =   13
      ToolTipText     =   "This will reload the information index and hence response will be slow."
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox pixSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pixDummy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1020
      Picture         =   "frmSearchInfo.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton radOR 
         Caption         =   "&OR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton radAND 
         Caption         =   "&AND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.ComboBox cbSearchFor 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtFullPath 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   5895
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "&Go"
      Default         =   -1  'True
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
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ListView lvSearchResults 
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "imlColumnHeaderIcons"
      SmallIcons      =   "imlColumnHeaderIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   9596
      EndProperty
   End
   Begin MSComctlLib.ImageList imlColumnHeaderIcons 
      Left            =   2400
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchInfo.frx":048C
            Key             =   "Dummy"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search For:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblSearchSum 
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note: To search for more than one word, separate each word by a space. e.g. Use 'VB PDF' to search for VB as well as PDFs."
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   10830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Full Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   825
   End
   Begin VB.Menu mOpenFile 
      Caption         =   "Open"
      Begin VB.Menu mOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mOpenWith 
         Caption         =   "Open File &With"
      End
      Begin VB.Menu mSpe1 
         Caption         =   "-"
      End
      Begin VB.Menu mOpenParent 
         Caption         =   "Open &Parent Folder"
      End
   End
End
Attribute VB_Name = "frmSearchInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###################################################################################
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'###################################################################################
Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18
'###################################################################################
Private Type POINTAPI
    x As Long
    y As Long
End Type
'###################################################################################
Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type
'###################################################################################
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
'###################################################################################
Private Const SW_MAXIMIZE = 3 'to maximize the window after it is opened with ShellExecute
'###################################################################################
'to show Open With Dialog box
Private Declare Sub OpenAs Lib "shell32.dll" Alias "OpenAs_RunDLLA" _
(ByVal hwnd As Long, ByVal res1 As Long, ByVal FileName As String, ByVal res2 As Long)
'###################################################################################
Private Sub btnGo_Click()
  search
End Sub
'###################################################################################
Private Sub btnRefresh_Click()
  search True
End Sub
'###################################################################################
Private Sub cbSearchFor_Click()
  search
End Sub
'###################################################################################
Private Sub Form_Load()
  Me.mOpenFile.Visible = False
  Me.lblSearchSum.Caption = ""
  ListHeaders
  Me.lvSearchResults.Sorted = True
  Me.lvSearchResults.SortKey = 2 'type
  Me.lvSearchResults.ColumnHeaders.item(3).Icon = ASCENDING 'type
  InitializeImageList
  search
End Sub
'###################################################################################
Private Function InitializeImageList() As Boolean
  On Local Error GoTo InitializeError
      Set Me.lvSearchResults.SmallIcons = Nothing
      ImageList1.ListImages.Clear
      ImageList1.ListImages.Add , "dummy", pixDummy.Picture
      Set Me.lvSearchResults.SmallIcons = Me.ImageList1
      InitializeImageList = True
    Exit Function
InitializeError:
  InitializeImageList = False
End Function
'###################################################################################
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
On Error GoTo Hell
  Me.lvSearchResults.Width = Me.Width - 700
  Me.lvSearchResults.Height = Me.Height - 2000
  Me.txtFullPath.Width = Me.Width - 3100
  Me.lblNote.Width = Me.lvSearchResults.Width
Hell:
  
End Sub
'###################################################################################
Private Sub lvSearchResults_Click()
  If Me.lvSearchResults.ListItems.Count > 0 Then
    Dim t_info As InfoBean
    Set t_info = Me.lvSearchResults.SelectedItem.Tag
    Me.txtFullPath.Text = t_info.m_fullpath
  End If
End Sub
'###################################################################################
'adds the text to the combo box "cbSearchFor"
Private Sub addToSearchHistory(sText As String)
  If Len(Trim(sText)) = 0 Then Exit Sub
  Dim i As Integer
  Dim blnFound As Boolean
  blnFound = False
  For i = 1 To Me.cbSearchFor.ListCount
    If UCase(sText) = UCase(Me.cbSearchFor.List(i)) Then
      blnFound = True
      Exit For
    End If
  Next
  If Not blnFound Then
    Me.cbSearchFor.AddItem (Me.cbSearchFor.Text)
  End If
End Sub
'###################################################################################
Public Sub search(Optional blnRefresh As Boolean = False)
  Dim infoIndex As Collection
  Set infoIndex = LoadInfoMod.getIndex(blnRefresh)
  Me.txtFullPath.Text = ""
  addToSearchHistory Me.cbSearchFor.Text
  
  Dim bIsAndOpr As Boolean
  bIsAndOpr = Me.radAND.Value
  
  Me.lvSearchResults.ListItems.Clear
  Dim lstItem As ListItem
  Dim fs As New FileSystemObject
  Dim i As Integer
  Dim arrSearchCriteria As Variant
  For i = 1 To infoIndex.Count
    Dim isMatching As Boolean
    isMatching = bIsAndOpr
    If (Len(Me.cbSearchFor.Text) = 0) Then
      isMatching = True
    Else
      Dim j As Integer
      arrSearchCriteria = Split(Me.cbSearchFor.Text, " ")
      For j = 0 To UBound(arrSearchCriteria)
        If bIsAndOpr Then
          If InStr(1, UCase(infoIndex.item(i)), UCase(arrSearchCriteria(j))) = 0 Then
            isMatching = False
            Exit For
          End If
        Else
          If InStr(1, UCase(infoIndex.item(i)), UCase(arrSearchCriteria(j))) > 0 Then
            isMatching = True
            Exit For
          End If
        End If
      Next
    End If
    'If InStr(1, UCase(infoIndex.Item(i)), UCase(Me.cbSearchFor.Text)) > 0 Then
    If (isMatching) Then
      'add to list
      Dim t_info As InfoBean
      Set t_info = New InfoBean
      Set lstItem = addFileToList(infoIndex.item(i))
      If (Not lstItem Is Nothing) Then
        t_info.init (infoIndex.item(i))
        If fs.FolderExists(infoIndex.item(i)) Then
          lstItem.SubItems(2) = "File Folder"
        Else
          If fs.FileExists(infoIndex.item(i)) Then
            Dim file As file
            Set file = fs.GetFile(infoIndex.item(i))
            lstItem.SubItems(2) = file.Type
          Else
            lstItem.SubItems(2) = "File"
          End If
        End If
        lstItem.SubItems(3) = t_info.getParentFolder
        Set lstItem.Tag = t_info
      Else
'        Dim s As String
'        s = infoIndex.item(i) & " not found."
      End If
    End If
  Next
  Me.lblSearchSum.Caption = "Count: " & Me.lvSearchResults.ListItems.Count
  Set fs = Nothing
End Sub
'###################################################################################
Public Function addFileToList(sFilePath As String) As ListItem
  Dim hFile As Long
  Dim WFD As WIN32_FIND_DATA
  hFile& = FindFirstFile(sFilePath, WFD)
  Dim lstItem As ListItem
  If hFile& > 0 Then
    Set lstItem = vbAddFileItemView(WFD, sFilePath)
  End If
  FindClose hFile
  Set addFileToList = lstItem
End Function
'###################################################################################
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
End Sub
'###################################################################################
Private Sub lvSearchResults_DblClick()
  'open item
  If Me.lvSearchResults.ListItems.Count > 0 Then
    Dim t_info As InfoBean
    Set t_info = Me.lvSearchResults.SelectedItem.Tag
    OpenFileOrFolder t_info
  End If
End Sub
'###################################################################################
Public Sub OpenFileOrFolder(ByRef p_info As InfoBean)
  Dim fs As New FileSystemObject
  If (fs.FileExists(p_info.m_fullpath)) Then
    'open file
    lngErr = ShellExecute(0, "OPEN", p_info.m_fullpath, "", "", SW_MAXIMIZE)
'    lngErr = ShellExecute(0, "open", Environ("windir") & "\notepad.exe", _
'                                  t_info.m_fullpath, Environ("windir"), 0)
'    If (lngErr <> 42 And lngErr <> 0 And lngErr <> 33) Then
    If (lngErr <= 32) Then 'error
      'show Choose Program dialog
      Call OpenAs(0, 0, p_info.m_fullpath, 0)
'      'open the folder containing the file instead
'      MsgBox "Unable to open the file. " & _
'              "Hence opening the folder containing the file.", vbExclamation, getApplnName
'      Dim sFolderPath As String
'      Dim index As String
'      If (fs.FolderExists(p_info.getParentFolder)) Then 'open folder
'        OpenSinglePaneExplorer p_info.getParentFolder, True
'      Else
'        MsgBox "Invalid path: '" & p_info.m_fullpath & "'", vbCritical, getApplnName
'      End If
    End If
  ElseIf (fs.FolderExists(p_info.m_fullpath)) Then 'open folder
    OpenSinglePaneExplorer p_info.m_fullpath
  End If
End Sub
'###################################################################################

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
'###################################################################################
Private Sub lvSearchResults_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim t_info As InfoBean
  Set t_info = Me.lvSearchResults.SelectedItem.Tag
  Me.txtFullPath.Text = t_info.m_fullpath
End Sub
'###################################################################################
Private Sub lvSearchResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = MouseButtonConstants.vbRightButton Then
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim t_info As InfoBean
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(Me.lvSearchResults.hwnd, LVM_HITTEST, 0, lvhti) + 1
    If (lItemIndex <> 0) Then
      Me.lvSearchResults.SelectedItem = Me.lvSearchResults.ListItems.item(lItemIndex)
      Set t_info = Me.lvSearchResults.SelectedItem.Tag
      Dim fs As New FileSystemObject
      If fs.FolderExists(t_info.m_fullpath) Then
        mOpenWith.Enabled = False
      Else
        mOpenWith.Enabled = True
      End If
      Set fs = Nothing
      Me.PopupMenu Me.mOpenFile
    End If
  End If
End Sub
'###################################################################################
Private Sub lvSearchResults_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim t_info As InfoBean
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(Me.lvSearchResults.hwnd, LVM_HITTEST, 0, lvhti) + 1
    If (lItemIndex <> 0) Then
      Me.lvSearchResults.SelectedItem = Me.lvSearchResults.ListItems.item(lItemIndex)
      Set t_info = Me.lvSearchResults.SelectedItem.Tag
      Me.txtFullPath.Text = t_info.m_fullpath
    End If
End Sub
'###################################################################################
Private Sub mOpen_Click()
  If (Me.lvSearchResults.ListItems.Count > 0) Then
    Dim t_info As InfoBean
    Set t_info = Me.lvSearchResults.SelectedItem.Tag
    OpenFileOrFolder t_info
  End If
End Sub
'###################################################################################
Private Sub mOpenParent_Click()
  If (Me.lvSearchResults.ListItems.Count > 0) Then
    Dim t_info As InfoBean
    Set t_info = Me.lvSearchResults.SelectedItem.Tag
    OpenSinglePaneExplorer t_info.getParentFolder, True
  End If
End Sub
'###################################################################################
Private Sub mOpenWith_Click()
   If (Me.lvSearchResults.ListItems.Count > 0) Then
    Dim t_info As InfoBean
    Set t_info = Me.lvSearchResults.SelectedItem.Tag
    Call OpenAs(0, 0, t_info.m_fullpath, 0)
  End If
End Sub
'###################################################################################
Private Sub radAND_Click()
  If Len(Trim(Me.cbSearchFor.Text)) > 0 Then
    search
  End If
End Sub
'###################################################################################
Private Sub radOR_Click()
  If Len(Trim(Me.cbSearchFor.Text)) > 0 Then
    search
  End If
End Sub
'###################################################################################
Private Function vbAddFileItemView(WFD As WIN32_FIND_DATA, sFilePath As String) As ListItem
    Dim sFileName As String
    Dim ListImgKey As String
    Dim fType As String
    
    sFileName = TrimNull(WFD.cFileName)
    
    If sFileName <> "." And sFileName <> ".." Then

      Dim r As Long
      Dim tExeType As Long
      Dim itmX As ListItem

      Dim hImgSmall As Long
      Dim hExeType As Long
      Dim imgX As ListImage
      
On Local Error GoTo AddFileItemViewError
      hImgSmall& = SHGetFileInfo(sFilePath, _
                       0&, shinfo, Len(shinfo), _
                       BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
      fType$ = TrimNull(shinfo.szTypeName)
      If Len(fType) = 0 Then
        fType$ = Mid(sFileName, Len(sFileName) - 3, 3) 'assign extn of the file
      End If
      ListImgKey = fType
      If fType = "application" Or fType = "shortcut" Then
        If fType = "application" Then
          tExeType = SHGetFileInfo(fPath & sFileName, _
                          0&, shinfo, Len(shinfo), SHGFI_EXETYPE)
          hExeType = HiWord(tExeType)
        End If
        If hExeType > 0 Or fType = "shortcut" Then
           r = vbAddFileItemIcon(hImgSmall)
           Set imgX = ImageList1.ListImages.Add(, sFileName, pixSmall.Picture)
           ListImgKey = sFileName
        Else
           ListImgKey = "DOSExeIcon"
           If DOSExeIconLoaded = False Then
              r = vbAddFileItemIcon(hImgSmall)
              Set imgX = ImageList1.ListImages.Add(, ListImgKey, pixSmall.Picture)
              DOSExeIconLoaded = True
           End If
        End If
      End If
      
      Set itmX = Me.lvSearchResults.ListItems.Add(, , sFileName)
      itmX.SmallIcon = ImageList1.ListImages(ListImgKey).Key
      If (fType = "File Folder") Then
        itmX.SubItems(1) = ""
      Else
        itmX.SubItems(1) = vbGetFileSizeKBStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
      End If
'      itmX.SubItems(2) = fType
'      itmX.SubItems(3) = vbGetFileDate(WFD.ftCreationTime)
      Set vbAddFileItemView = itmX
    End If
  Exit Function
AddFileItemViewError:
    If vbAddFileItemIcon(hImgSmall) Then
      Set imgX = ImageList1.ListImages.Add(, fType, pixSmall.Picture)
    End If
  Resume
End Function
'###################################################################################
Private Function vbAddFileItemIcon(hImgSmall&) As Long
    Dim r As Long
    pixSmall.Picture = LoadPicture()
    r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, pixSmall.hDC, 0, 0, ILD_TRANSPARENT)
    pixSmall.Picture = pixSmall.Image
    vbAddFileItemIcon& = hImgSmall&
End Function
'###################################################################################
