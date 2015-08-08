Attribute VB_Name = "ExtractFileIconMod"
'###################################################################################
Option Explicit
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH As Long = 260
'###################################################################################
Public Type SHITEMID
  cb      As Long
  abID    As Byte
End Type
'###################################################################################
Public Type ITEMIDLIST
  mkid    As SHITEMID
End Type
'###################################################################################
Public Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type
'###################################################################################
Public Declare Function SHGetPathFromIDList Lib "Shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long
'###################################################################################
Public Declare Function SHBrowseForFolder Lib "Shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long
'###################################################################################
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
'###################################################################################
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'To the Constant declarations add:
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Public Const SHGFI_LARGEICON = &H0        'large icon
Public Const SHGFI_SMALLICON = &H1        'small icon
Public Const ILD_TRANSPARENT = &H1        'display transparent
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
'###################################################################################
Public Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type
'###################################################################################
Public shinfo As SHFILEINFO
'###################################################################################
Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type
'###################################################################################
Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Integer
End Type
'###################################################################################
Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type
'###################################################################################
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'###################################################################################
Public Declare Function FindNextFile Lib "kernel32" _
  Alias "FindNextFileA" _
  (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'###################################################################################
Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
'###################################################################################
Public Declare Function FileTimeToSystemTime Lib "kernel32" _
  (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'###################################################################################
Public Declare Function UpdateWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
'###################################################################################
Public Declare Function ImageList_Draw Lib "comctl32" _
  (ByVal himl&, _
   ByVal i&, _
   ByVal hDCDest&, _
   ByVal x&, _
   ByVal y&, _
   ByVal flags&) As Long
'###################################################################################
Public Declare Function SHGetFileInfo Lib "Shell32" _
   Alias "SHGetFileInfoA" _
  (ByVal pszPath As String, _
   ByVal dwFileAttributes As Long, _
   psfi As SHFILEINFO, _
   ByVal cbSizeFileInfo As Long, _
   ByVal uFlags As Long) As Long
'###################################################################################
Public Function TrimNull(item As String) As String
    Dim pos As Integer
    pos = InStr(item, Chr$(0))
    If pos Then item = Left$(item, pos - 1)
    TrimNull = item
End Function
'###################################################################################
Public Function HiWord(dw As Long) As Integer
    If dw And &H80000000 Then
       HiWord = (dw \ 65535) - 1
    Else
       HiWord = dw \ 65535
    End If
End Function
'###################################################################################
Public Function vbGetFileSizeKBStr(fsize As Long) As String
    vbGetFileSizeKBStr = Format$(((fsize) / 1000) + 0.5, "#,###,###") & " KB"
End Function
'###################################################################################
Public Function vbGetFileDate(CT As FILETIME) As String
    Dim ST As SYSTEMTIME
    Dim ds As Single
    If FileTimeToSystemTime(CT, ST) Then
       ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
       vbGetFileDate$ = Format$(ds, "Short Date")
    Else
       vbGetFileDate$ = ""
    End If
End Function
'###################################################################################

