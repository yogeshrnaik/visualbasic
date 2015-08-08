Attribute VB_Name = "FormatList"
Option Explicit

' Color Constants
Public Const vbViolet = &HFF8080
Public Const vbVioletBright = &HFFC0C0
Public Const vbForestGreen = &H228B22


Public Const vbGray = &HE0E0E0
Public Const vbLightBlue = &HFFD3A4
Public Const vbLightGreen = &HABFCBD
Public Const vbGreenLemon = &HB3FFBE
Public Const vbYellowBright = &HC0FFFF
Public Const vbOrange = &H2CCDFC

Public Sub SetListViewColor(pCtrlListView As ListView, _
                            pCtrlPictureBox As PictureBox, _
                            Color1 As Long, _
                            Color2 As Long)

On Error GoTo SetListViewColor_Error

    Dim iLineHeight As Long
    Dim iBarHeight  As Long
    Dim lBarWidth   As Long
    Dim lColor1     As Long
    Dim lColor2     As Long
 
    lColor1 = Color1
    lColor2 = Color2
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        pCtrlPictureBox.Cls
        
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        pCtrlListView.PictureAlignment = lvwTile
        pCtrlPictureBox.Font = pCtrlListView.Font
        pCtrlPictureBox.Top = pCtrlListView.Top
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size + 2.75
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        'paint the two bars of color
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
        
        pCtrlPictureBox.AutoSize = True
        'set the pCtrlListView picture to the
        'pCtrlPictureBox image
        pCtrlListView.Picture = pCtrlPictureBox.Image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
SetListViewColor_Error:
    'clear pCtrlListView's picture and then exit
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
End Sub


