Attribute VB_Name = "GeneralMod"
'###################################################################################
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type
'###################################################################################
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long
'###################################################################################
Private Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" _
    (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
    ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
'###################################################################################
Private Const MAX_PATH = 260
'###################################################################################
Public Const NO_ORDER = 0
Public Const DESCENDING = 1
Public Const ASCENDING = 2
'###################################################################################
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
'###################################################################################
Private Const SHGFI_ICON = &H100
Private Const SHGFI_OPENICON = &H2
Private Const SHGFI_SELECTED = &H10000
'###################################################################################
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_LARGEICON = &H0
'###################################################################################
Public Enum mbIconSizeConstants
    mbLargeIcon = SHGFI_LARGEICON           '32x32 icon
    mbSmallIcon = SHGFI_SMALLICON           '16x16 icon
    mbShellSizeIcon = SHGFI_SHELLICONSIZE   'size used by the shell to display
End Enum                                    'the icons (for example 32x32 or
                                            ' 48x48)
'###################################################################################
Public Enum mbIconTypeConstants
    mbNormalIcon = SHGFI_ICON               'normal icon
    mbSelectedIcon = SHGFI_SELECTED         'the icon when it is selected
    mbOpenIcon = SHGFI_OPENICON             'the icon used for open folders
End Enum
'###################################################################################
' Returns the description of the specified file/folder (for example "Folder",
' "Executable file", "Bmp Image" and so on)
' Get the file/folder's associated icon
' NOTE: uses the IconToPicture function (you can find it elsewhere in the Code Bank)
Function GetFileIcon(ByVal sPath As String, Optional ByVal mbIconSize As _
    mbIconSizeConstants = mbLargeIcon, Optional ByVal mbIconType As _
    mbIconTypeConstants = mbNormalIcon) As StdPicture
    Dim FInfo As SHFILEINFO
    Dim lIconType As Long
    
    lIconType = mbIconSize Or mbIconType
    ' be sure that there is the mbNormalIcon too
    If mbIconType <> mbNormalIcon Then lIconType = lIconType Or mbNormalIcon
    ' retrieve the item's icon
    SHGetFileInfo sPath, 0, FInfo, Len(FInfo), lIconType
    ' convert the handle to a StdPicture
    Set GetFileIcon = IconToPicture(FInfo.hIcon)
End Function
'###################################################################################
Public Function getApplnName() As String
  getApplnName = "InfoQuest"
End Function
'###################################################################################
'OpenSinglePaneExplorer - Open a folder in a single pane Explorer window
' sRoot is the root folder to open
' bUpwardAllowed specifies whether the user will be allowed to navigate upward the root
' Example: OpenSinglePaneExplorer "D:\Webmaster", False
Public Sub OpenSinglePaneExplorer(Optional ByVal sRoot As String, _
    Optional ByVal bUpwardAllowed As Boolean = True)
    Shell "explorer /e," & IIf(bUpwardAllowed, "", "/root,") & sRoot, vbMaximizedFocus
End Sub
'###################################################################################
'IconToPicture - Convert an icon handle to a Picture object
' Convert an icon handle to a Picture object
Function IconToPicture(ByVal hIcon As Long) As Picture
    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    pic.pictType = vbPicTypeIcon
    pic.hIcon = hIcon
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, IconToPicture
End Function




