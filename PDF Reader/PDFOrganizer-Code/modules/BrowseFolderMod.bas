Attribute VB_Name = "BrowseFolderMod"
'#####################################################################################
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declarations for Browse Folder dialog
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH As Long = 260

Private Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)
'#####################################################################################
Public Function displayFolderDialog(parentForm As Form) As String
  Dim bi As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Long
  
 'Fill the BROWSEINFO structure with the
 'needed data. To accommodate comments, the
 'With/End With syntax has not been used, though
 'it should be your 'final' version.

  With bi
    'hwnd of the window that receives messages
    'from the call. Can be your application
    'or the handle from GetDesktopWindow()
    .hOwner = parentForm.hwnd
    
    'pointer to the item identifier list specifying
    'the location of the "root" folder to browse from.
    'If NULL, the desktop folder is used.
    '.pidlRoot = 0&
    
    'message to be displayed in the Browse dialog
    .lpszTitle = "Select your Windows\System\ directory"
    
    'the type of folder to return.
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  'show the Browse Dialog
   pidl = SHBrowseForFolder(bi)
 
  'the dialog has closed, so parse & display the
  'user's returned folder selection contained in pidl
   path = Space$(MAX_PATH)
   If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
      pos = InStr(path, Chr$(0))
      path = Left(path, pos - 1)
    Else
      path = ""
   End If
  Call CoTaskMemFree(pidl)
  displayFolderDialog = path
End Function
'#####################################################################################

'#####################################################################################

'#####################################################################################

'#####################################################################################
