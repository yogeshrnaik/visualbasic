VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MSExcel - A class for writing Excel spreadsheets
'***********************************************************
'* Module Name: MSExcel
'* Author:      Steve Miller
'* Date:        11/22/98
'* Description: Encapsulates and eases the chore of writing to an Excel
' spreadsheet.
'*
'* IMPORTANT: requires that you have a reference to Excel type library
'*
'* Example:     Here is an example of how to use this class:
'*   Dim objXL As New ttExcel.Application                   'Instantiate the
' object
'*   Call objXL.OpenExcelFile(FileName:="c:\SMTest.xls",
'  '*        WorkSheetName:="MyWorksheet", '*        AppendDayofWk:=True)
'                        'Open the Excel file
'*   Call objXL.StoreExcelRow(A:="Hello", B:="World")       'Write a row to the
' Excel file
'*   Call objXL.StoreExcelBlankRows(NbrRows:=10)            'Skip blank rows in
' the Excel file
'*   Call objXL.StoreExcelRow(A:="Skipped 10 rows")
'*   Call objXL.SetExcelCellWidthAutoFit                    'Auto size the
' column widths
'*   Call objXL.SetExcelCellWidth(Cell:="A", Width:=25)     'Manually set a
' column width
'*   Call objXL.CloseExcelFile                              'Close the Excel
' file
'*   Set objXL = Nothing                                    'Cleanup (release
' memory)
'***********************************************************

Option Compare Text
Option Base 1                    'Start arrays at element 1 instead of 0



'Constants
Private Const m_constPgm As String = "MSExcel."
Private Const m_lngExcelLabelNotFound = 1004

'Objects
Private m_objWorkbook             As Workbook
Private m_objWorksheet            As Worksheet
Private m_objXL                   As Excel.Application
Private m_objPageSetup            As Excel.PageSetup
Private m_lxPageOrientation       As XlPageOrientation

'Strings
Private m_strExcelFileName          As String

'Longs
Private m_lngExcelRow               As Long     'SM 6/24/98 Used for Excel Row
                                                ' Number
'*--------------------------------------------------------------
'* StoreExcelRow
'* - Common routine used for writing a row to Excel.  The nice thing about this
'*   routine is that you do not have to keep up with the row number.  It will
'*   automatically increment the row number so that you don't have to keep up
' with it.
'* - Inputs:
'*   A - Z represent columns in Excel.  For example, to write information in
'*         column A, you assign information to A.
'*   FontBold - True if you want the entire row to be bold (used for headers/
' titles)
'* - Examples:
'*   Call StoreExcelRow(A:="Column A", B:="Column B", FontBold:=True
'*   Call StoreExcelRow(A:="12344.33", B:="My Text!", FontBold:=False
'*--------------------------------------------------------------
Public Sub StoreExcelRow(Optional A As Variant, Optional B As Variant, Optional C As Variant, Optional D As Variant, Optional E As Variant, _
    Optional F As Variant, Optional G As Variant, Optional H As Variant, Optional i As Variant, Optional J As Variant, Optional K As Variant, _
    Optional L As Variant, Optional M As Variant, Optional n As Variant, Optional O As Variant, Optional P As Variant, Optional Q As Variant, _
    Optional R As Variant, Optional S As Variant, Optional T As Variant, Optional U As Variant, Optional V As Variant, Optional W As Variant, _
    Optional x As Variant, Optional Y As Variant, Optional Z As Variant, Optional FontBold As Boolean = False, Optional FontSize As Integer, _
    Optional FontItalic As Boolean = False, Optional FontOutline As Boolean = False, Optional FontColor As Variant, Optional ALabel As Variant, _
    Optional BLabel As Variant, Optional CLabel As Variant, Optional DLabel As Variant, Optional ELabel As Variant, _
    Optional FLabel As Variant, Optional GLabel As Variant, Optional HLabel As Variant, Optional ILabel As Variant, _
    Optional JLabel As Variant, Optional KLabel As Variant, Optional LLabel As Variant, Optional MLabel As Variant, _
    Optional nLabel As Variant, Optional OLabel As Variant, Optional PLabel As Variant, Optional QLabel As Variant, _
    Optional RLabel As Variant, Optional SLabel As Variant, Optional TLabel As Variant, Optional ULabel As Variant, _
    Optional VLabel As Variant, Optional WLabel As Variant, Optional XLabel As Variant, Optional YLabel As Variant, _
    Optional ZLabel As Variant, Optional BorderHeader As Boolean = False, Optional BorderAround = False)
On Error GoTo ErrorHandler
   Const constErrRoutine As String = "StoreExcelRow"
   Dim strBeginCell As String
   Dim strEndCell As String
   Dim intCounter As Integer
   Dim intBegin As Integer
   Dim intEnd As Integer
   Const constCells As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   Const constBeige As Variant = &HC0FFFF
   Const constDarkBlue As Variant = &H800000

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   m_lngExcelRow = m_lngExcelRow + 1
   If Not IsMissing(A) Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow))).Value = CStr(A)
      strBeginCell = "A"
   End If

   If Not IsMissing(B) Then
      m_objWorksheet.Range("B" & Trim(Str(m_lngExcelRow))).Value = CStr(B)
      If strBeginCell = "" Then
         strBeginCell = "B"
      Else
         strEndCell = "B"
      End If
   End If

   If Not IsMissing(C) Then
      m_objWorksheet.Range("C" & Trim(Str(m_lngExcelRow))).Value = CStr(C)
      If strBeginCell = "" Then
         strBeginCell = "C"
      Else
         strEndCell = "C"
      End If
   End If

   If Not IsMissing(D) Then
      m_objWorksheet.Range("D" & Trim(Str(m_lngExcelRow))).Value = CStr(D)
      If strBeginCell = "" Then
         strBeginCell = "D"
      Else
         strEndCell = "D"
      End If
   End If


   If Not IsMissing(E) Then
      m_objWorksheet.Range("E" & Trim(Str(m_lngExcelRow))).Value = CStr(E)
      If strBeginCell = "" Then
         strBeginCell = "E"
      Else
         strEndCell = "E"
      End If
   End If

   If Not IsMissing(F) Then
      m_objWorksheet.Range("F" & Trim(Str(m_lngExcelRow))).Value = CStr(F)
      If strBeginCell = "" Then
         strBeginCell = "F"
      Else
         strEndCell = "F"
      End If
   End If

   If Not IsMissing(G) Then
      m_objWorksheet.Range("G" & Trim(Str(m_lngExcelRow))).Value = CStr(G)
      If strBeginCell = "" Then
         strBeginCell = "G"
      Else
         strEndCell = "G"
      End If
   End If

   If Not IsMissing(H) Then
      m_objWorksheet.Range("H" & Trim(Str(m_lngExcelRow))).Value = CStr(H)
      If strBeginCell = "" Then
         strBeginCell = "H"
      Else
         strEndCell = "H"
      End If
   End If

   If Not IsMissing(i) Then
      m_objWorksheet.Range("I" & Trim(Str(m_lngExcelRow))).Value = CStr(i)
      If strBeginCell = "" Then
         strBeginCell = "I"
      Else
         strEndCell = "I"
      End If
   End If

   If Not IsMissing(J) Then
      m_objWorksheet.Range("J" & Trim(Str(m_lngExcelRow))).Value = CStr(J)
      If strBeginCell = "" Then
         strBeginCell = "J"
      Else
         strEndCell = "J"
      End If
   End If

   If Not IsMissing(K) Then
      m_objWorksheet.Range("K" & Trim(Str(m_lngExcelRow))).Value = CStr(K)
      If strBeginCell = "" Then
         strBeginCell = "K"
      Else
         strEndCell = "K"
      End If
   End If

   If Not IsMissing(L) Then
      m_objWorksheet.Range("L" & Trim(Str(m_lngExcelRow))).Value = CStr(L)
      If strBeginCell = "" Then
         strBeginCell = "L"
      Else
         strEndCell = "L"
      End If
   End If

   If Not IsMissing(M) Then
      m_objWorksheet.Range("M" & Trim(Str(m_lngExcelRow))).Value = CStr(M)
      If strBeginCell = "" Then
         strBeginCell = "M"
      Else
         strEndCell = "M"
      End If
   End If

   If Not IsMissing(n) Then
      m_objWorksheet.Range("N" & Trim(Str(m_lngExcelRow))).Value = CStr(n)
      If strBeginCell = "" Then
         strBeginCell = "N"
      Else
         strEndCell = "N"
      End If
   End If

   If Not IsMissing(O) Then
      m_objWorksheet.Range("O" & Trim(Str(m_lngExcelRow))).Value = CStr(O)
      If strBeginCell = "" Then
         strBeginCell = "O"
      Else
         strEndCell = "O"
      End If
   End If

   If Not IsMissing(P) Then
      m_objWorksheet.Range("P" & Trim(Str(m_lngExcelRow))).Value = CStr(P)
      If strBeginCell = "" Then
         strBeginCell = "P"
      Else
         strEndCell = "P"
      End If
   End If

   If Not IsMissing(Q) Then
      m_objWorksheet.Range("Q" & Trim(Str(m_lngExcelRow))).Value = CStr(Q)
      If strBeginCell = "" Then
         strBeginCell = "Q"
      Else
         strEndCell = "Q"
      End If
   End If

   If Not IsMissing(R) Then
      m_objWorksheet.Range("R" & Trim(Str(m_lngExcelRow))).Value = CStr(R)
      If strBeginCell = "" Then
         strBeginCell = "R"
      Else
         strEndCell = "R"
      End If
   End If

   If Not IsMissing(S) Then
      m_objWorksheet.Range("S" & Trim(Str(m_lngExcelRow))).Value = CStr(S)
      If strBeginCell = "" Then
         strBeginCell = "S"
      Else
         strEndCell = "S"
      End If
   End If

   If Not IsMissing(T) Then
      m_objWorksheet.Range("T" & Trim(Str(m_lngExcelRow))).Value = CStr(T)
      If strBeginCell = "" Then
         strBeginCell = "T"
      Else
         strEndCell = "T"
      End If
   End If

   If Not IsMissing(U) Then
      m_objWorksheet.Range("U" & Trim(Str(m_lngExcelRow))).Value = CStr(U)
      If strBeginCell = "" Then
         strBeginCell = "U"
      Else
         strEndCell = "U"
      End If
   End If

   If Not IsMissing(V) Then
      m_objWorksheet.Range("V" & Trim(Str(m_lngExcelRow))).Value = CStr(V)
      If strBeginCell = "" Then
         strBeginCell = "V"
      Else
         strEndCell = "V"
      End If
   End If

   If Not IsMissing(W) Then
      m_objWorksheet.Range("W" & Trim(Str(m_lngExcelRow))).Value = CStr(W)
      If strBeginCell = "" Then
         strBeginCell = "W"
      Else
         strEndCell = "W"
      End If
   End If

   If Not IsMissing(x) Then
      m_objWorksheet.Range("X" & Trim(Str(m_lngExcelRow))).Value = CStr(x)
      If strBeginCell = "" Then
         strBeginCell = "X"
      Else
         strEndCell = "X"
      End If
   End If

   If Not IsMissing(Y) Then
      m_objWorksheet.Range("Y" & Trim(Str(m_lngExcelRow))).Value = CStr(Y)
      If strBeginCell = "" Then
         strBeginCell = "Y"
      Else
         strEndCell = "Y"
      End If
   End If

   If Not IsMissing(Z) Then
      m_objWorksheet.Range("Z" & Trim(Str(m_lngExcelRow))).Value = CStr(Z)
      If strBeginCell = "" Then
         strBeginCell = "Z"
      Else
         strEndCell = "Z"
      End If
   End If

   If Not IsMissing(ALabel) Then m_objWorksheet.Range("A" & Trim(Str _
       (m_lngExcelRow))).Name = ALabel

   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Bold = FontBold
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Italic = FontItalic
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.OutlineFont = FontOutline

   If IsMissing(FontSize) Or FontSize = 0 Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = 10
   Else
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = FontSize
   End If

   If Not IsMissing(FontColor) Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Color = FontColor
   End If

   If strBeginCell = "" Then strBeginCell = "A"
   If strEndCell = "" Then strEndCell = "Z"

   If BorderHeader Then
      'Show as a heading by making the background dark blue and the font white
      intBegin = InStr(1, constCells, strBeginCell)
      intEnd = InStr(1, constCells, strEndCell)
      For intCounter = intBegin To intEnd
         Call DrawCellBorder(Cell:=Mid$(constCells, intCounter, 1), _
             InteriorColor:=constDarkBlue)
      Next intCounter
      m_objWorksheet.Range(strBeginCell & Trim(Str(m_lngExcelRow)) & ":" & _
          strEndCell & Trim(Str(m_lngExcelRow))).Font.Color = vbWhite
   End If

   If BorderAround Then
      'Put a thin border around the area and make background beige
      intBegin = InStr(1, constCells, strBeginCell)
      intEnd = InStr(1, constCells, strEndCell)
      For intCounter = intBegin To intEnd
         Call DrawCellBorder(Cell:=Mid$(constCells, intCounter, 1), _
             InteriorColor:=constBeige)
      Next intCounter
   End If

ExitMe:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, m_constPgm & ".StoreExcelRow", Err.Description
    Resume ExitMe

   End Sub

'*--------------------------------------------------------------
'* DrawCellBorder
'* - Draws a border around the cell
'*--------------------------------------------------------------
Private Sub DrawCellBorder(Cell As String, InteriorColor As Variant)

   With m_objWorksheet.Range(Cell & Trim(Str(m_lngExcelRow)) & ":" & Cell & _
       Trim(Str(m_lngExcelRow)))
      .Cells.Interior.Color = InteriorColor
      .BorderAround Color:=vbBlack
   End With

End Sub

'*--------------------------------------------------------------
'* CreateExcelFromADORecordset
'* - Writes the rows from the ADO Recordset
'* - Inputs:
'*   rst = ADO Recordset
'*   Title = Descriptive Title on the spreadsheet
'*   FontColor = Color of the font (like vbBlue, vbBlack, etc)
'*   FontSize = Size of the font (like 10 pt)
'* - Example:     Call CreateExcelFromADORecordset(rst,vbBlue,10)
'*--------------------------------------------------------------

Public Sub CreateExcelFromADORecordset(rst As Object, _
    Optional strTitle As Variant, Optional FontColor As Variant, _
    Optional FontSize As Integer)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "CreateExcelFromADORecordset"
   Const constCell As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   Dim i As Integer
   Dim sCell As String
   Dim vValue As Variant
   m_lngExcelRow = m_lngExcelRow + 1

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   If Not IsMissing(strTitle) Then
      StoreExcelBlankRows (4)
   End If

   If IsMissing(FontSize) Or FontSize = 0 Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = 10
   Else
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = FontSize
   End If

   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Bold = True
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Color = vbBlue

   'First write the column names
   For i = 0 To rst.Fields.Count - 1
      sCell = Mid(constCell, i + 1, 1) & Trim$(Str(m_lngExcelRow))
      vValue = ProperCase(rst.Fields.Item(i).Name)
             'Must put it in a variant to keep from getting error
      If IsNumeric(vValue) Then
         m_objWorksheet.Range(sCell).Value = Str(vValue)
      Else
         m_objWorksheet.Range(sCell).Value = vValue
      End If
   Next i

   m_lngExcelRow = m_lngExcelRow + 1

   If Not IsMissing(FontColor) Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Color = FontColor
   End If

   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Bold = False
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Color = vbBlack

   'Now write each row
   rst.MoveFirst
   Do Until rst.EOF
      For i = 0 To rst.Fields.Count - 1
         sCell = Mid$(constCell, i + 1, 1) & Trim$(Str(m_lngExcelRow))
         vValue = rst.Fields.Item(i).Value         'Must put it in a variant to
                                                   ' keep from getting error
         If IsNumeric(vValue) Then
            m_objWorksheet.Range(sCell).Value = Str(vValue)
         Else
            m_objWorksheet.Range(sCell).Value = vValue
         End If
      Next i
      rst.MoveNext
      m_lngExcelRow = m_lngExcelRow + 1
   Loop

   SetExcelCellWidthAutoFit

   'Now write title if needed
   If Not IsMissing(strTitle) Then
      m_objWorksheet.Range("A1:Z1").Font.Bold = True
      m_objWorksheet.Range("A1:Z1").Font.Color = vbBlack
      m_objWorksheet.Range("A1:Z1").Font.Size = 18
      m_objWorksheet.Range("A1") = strTitle
      m_objWorksheet.Range("A2:Z2").Font.Bold = True
      m_objWorksheet.Range("A2:Z2").Font.Color = vbBlack
      m_objWorksheet.Range("A2:Z2").Font.Size = 10
      m_objWorksheet.Range("A2") = "Run: " & Format$(Now, _
          "mmmm dd, yyyy hh:mm AMPM")
      m_objWorksheet.Range("A3:Z3").Font.Bold = True
      m_objWorksheet.Range("A3:Z3").Font.Color = vbBlack
      m_objWorksheet.Range("A3:Z3").Font.Size = 10
      m_objWorksheet.Range("A3") = rst.RecordCount & " rows listed below "
   End If

   'Disable error trap in case this is a forward-only cursor
   On Error Resume Next
   rst.MoveFirst

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* CreateExcelFromSSOLEDBGrid
'* - Writes the rows from the Data Widgets grid
'* - Inputs:
'*   grd = SSOLEDbGrid
'*   Title = Descriptive Title on the spreadsheet
'*   FontColor = Color of the font (like vbBlue, vbBlack, etc)
'*   FontSize = Size of the font (like 10 pt)
'* - Example:     Call CreateExcelFromSSOLEDBGrid(grd,vbBlue,10)
'*--------------------------------------------------------------

Public Sub CreateExcelFromSSOLEDBGrid(grd As Object, _
    Optional strTitle As Variant, Optional FontColor As Variant, _
    Optional FontSize As Integer)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "CreateExcelFromSSOLEDBGrid"
   Const constCell As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   Dim i As Integer
   Dim n As Integer
   Dim sCell As String
   Dim vValue As Variant
   Dim lngRows As Long
   Dim intCols As Integer

   grd.redraw = False
   m_lngExcelRow = m_lngExcelRow + 1

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromSSOLEDBGrid", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   If Not IsMissing(strTitle) Then
      StoreExcelBlankRows (4)
   End If

   If IsMissing(FontSize) Or FontSize = 0 Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = 10
   Else
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Size = FontSize
   End If

   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Bold = True
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Color = vbBlue

   'First write the column names
   intCols = grd.cols - 1

   'Only allow 26 columns (columns A-Z in spreadsheet)
   If intCols > 26 Then
    intCols = 25
   End If

   For i = 0 To intCols
      sCell = Mid(constCell, i + 1, 1) & Trim$(Str(m_lngExcelRow))
      vValue = grd.Columns(i).Caption
      If IsNumeric(vValue) Then
         m_objWorksheet.Range(sCell).Value = Str(vValue)
      Else
         m_objWorksheet.Range(sCell).Value = vValue
      End If
   Next i

   m_lngExcelRow = m_lngExcelRow + 1

   If Not IsMissing(FontColor) Then
      m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
          (m_lngExcelRow))).Font.Color = FontColor
   End If

   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Bold = False
   m_objWorksheet.Range("A" & Trim(Str(m_lngExcelRow)) & ":Z" & Trim(Str _
       (m_lngExcelRow))).Font.Color = vbBlack

   'Now write each row
   grd.MoveFirst
   lngRows = grd.Rows
   For i = 1 To lngRows
      For n = 0 To intCols
         sCell = Mid$(constCell, n + 1, 1) & Trim$(Str(m_lngExcelRow))
         vValue = grd.Columns(n).text        'Must put it in a variant to keep
                                             ' from getting error
         If IsNumeric(vValue) Then
            m_objWorksheet.Range(sCell).Value = Str(vValue)
         Else
            m_objWorksheet.Range(sCell).Value = vValue
         End If
      Next n
      grd.MoveNext
      m_lngExcelRow = m_lngExcelRow + 1
   Next i

   grd.MoveFirst
   SetExcelCellWidthAutoFit

   'Now write title if needed
   If Not IsMissing(strTitle) Then
      m_objWorksheet.Range("A1:Z1").Font.Bold = True
      m_objWorksheet.Range("A1:Z1").Font.Color = vbBlack
      m_objWorksheet.Range("A1:Z1").Font.Size = 18
      m_objWorksheet.Range("A1") = strTitle
      m_objWorksheet.Range("A2:Z2").Font.Bold = True
      m_objWorksheet.Range("A2:Z2").Font.Color = vbBlack
      m_objWorksheet.Range("A2:Z2").Font.Size = 10
      m_objWorksheet.Range("A2") = "Run: " & Format$(Now, _
          "mmmm dd, yyyy hh:mm AMPM")
      m_objWorksheet.Range("A3:Z3").Font.Bold = True
      m_objWorksheet.Range("A3:Z3").Font.Color = vbBlack
      m_objWorksheet.Range("A3:Z3").Font.Size = 10
      m_objWorksheet.Range("A3") = lngRows & " rows listed below "
   End If

ExitMe:
   grd.redraw = True
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* ProperCase
'* - Capitalizes the first character of the word and any character following an
'*   underscore, hypen or space. All other characters will be lowercase
'*--------------------------------------------------------------
Private Function ProperCase(NewValue As String) As String

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "ProperCase"
   Dim i As Integer

   ProperCase = ""

   'Capitalize the first character of the word and any character following an
   ' underscore, hypen or space
   'All other characters will be lowercase
   For i = 1 To Len(NewValue)
      If i = 1 Then
         ProperCase = UCase(Mid$(NewValue, 1, 1))
      Else
         If Mid$(NewValue, i, 1) = "_" Or Mid$(NewValue, i, _
             1) = " " Or Mid$(NewValue, i, 1) = "-" Then
            ProperCase = ProperCase & Mid$(NewValue, i, 1)
            i = i + 1
            ProperCase = ProperCase & UCase(Mid$(NewValue, i, 1))
         Else
            ProperCase = ProperCase & LCase(Mid$(NewValue, i, 1))
         End If
      End If
   Next i

ExitMe:
   Exit Function

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Function

'*--------------------------------------------------------------
'* StoreExcelBlankRows
'* - Skips line(s) in the excel file to produce blank rows
'*--------------------------------------------------------------

Public Sub StoreExcelBlankRows(NbrRows As Integer)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "StoreExcelBlankRows"
   Dim i    As Integer

  If m_objWorksheet Is Nothing Then
     Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
         "You must first invoke the OpenExcelFile method before calling this " _
         & "method."
     GoTo ExitMe
  End If

   For i = 1 To NbrRows
      Call StoreExcelRow(A:="  ")
   Next i

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* SetExcelCellWidth
'* - Sets the width of the excel cell(s).
'*--------------------------------------------------------------

Public Sub SetExcelCellWidth(Cell As String, Width As Single)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "SetExcelCellWidth"

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   With m_objWorksheet
       .Cells.Columns(Cell).ColumnWidth = Width
   End With

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* SetExcelCellWidthAutoFit
'* - Sets the width of the excel cell(s).
'*--------------------------------------------------------------

Public Sub SetExcelCellWidthAutoFit()

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "SetExcelCellWidthAutoFit"

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   m_objWorksheet.Columns("A:Z").AutoFit
   m_objWorksheet.Range("A1", "Z" & m_lngExcelRow).Rows.AutoFit

ExitMe:

   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* WrapText
'* - Wraps the text of a set of columns
'*--------------------------------------------------------------

Public Sub WrapText(BeginCell As String, EndCell As String)

   m_objWorksheet.Range(BeginCell & "1", EndCell & m_lngExcelRow).WrapText = _
       True

End Sub

'*--------------------------------------------------------------
'* AutoFitRows
'* - Autofits the text of a set of columns
'*--------------------------------------------------------------

Public Sub AutoFitRows(BeginCell As String, EndCell As String)

   m_objWorksheet.Range(BeginCell & "1", EndCell & m_lngExcelRow).Rows.AutoFit

End Sub

'*--------------------------------------------------------------
'* AutoFitColumns
'* - Autofits the text of a set of columns
'*--------------------------------------------------------------

Public Sub AutoFitColumns(BeginCell As String, EndCell As String)

   m_objWorksheet.Columns(BeginCell & ":" & EndCell).AutoFit

End Sub

'*--------------------------------------------------------------
'* CreateBorder
'* - Sets the width of the excel cell(s).
'*--------------------------------------------------------------

Public Sub CreateBorder(Cell1 As String, Cell2 As Single)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "CreateBorder"

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   With m_objWorksheet
       .Range(Cell1, Cell2).BorderAround
   End With

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* CreateBorder
'* - Sets the width of the excel cell(s).
'*--------------------------------------------------------------

Public Sub CreateBorderHeader(Cell1 As String, Cell2 As Single, _
    Optional FontColor As Variant)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "CreateBorder"

   If IsMissing(FontColor) Then
      FontColor = &H800000          'Dark Blue
   End If

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   With m_objWorksheet
       .Range(Cell1, Cell2).BorderAround Color:=FontColor
       .Range(Cell1, Cell2).Font.Color = vbWhite
   End With

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* OpenExcelFile
'* -  Returns a fully qualified Excel file name based on the export_path in the
'*    INI file.  Deletes existing file if applicable.
'*--------------------------------------------------------------

Public Sub OpenExcelFile(FileName As String, Optional WorkSheetName As String = _
    "", Optional AppendDayofWk As Boolean = False, Optional PageOrientation As _
    XlPageOrientation = xlLandscape)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "OpenExcelFile"
   Dim sReturnString     As String
   Dim lRc               As Long

   m_strExcelFileName = FileName

   If WorkSheetName = "" Then
      WorkSheetName = "Sheet1"
   End If

   'Append the day of the week to the file name if necessary...
   If AppendDayofWk Then
      If Right$(m_strExcelFileName, 4) = ".xls" Then
         m_strExcelFileName = Left$(m_strExcelFileName, _
             Len(m_strExcelFileName) - 4) & Format$(Now, "dd") & ".xls"
      Else
         m_strExcelFileName = Left$(m_strExcelFileName, _
             Len(m_strExcelFileName)) & Format$(Now, "dd") & ".xls"
      End If
   End If

   If Right$(m_strExcelFileName, 4) <> ".xls" Then
      m_strExcelFileName = m_strExcelFileName & ".xls"
   End If

   If Exists(m_strExcelFileName) Then
      Kill m_strExcelFileName
   End If

   Set m_objXL = New Excel.Application
   Set m_objWorkbook = m_objXL.Workbooks.Add
   Set m_objWorksheet = m_objWorkbook.ActiveSheet
   m_objWorksheet.Activate

   'Set the worksheet name
   m_objWorksheet.Name = WorkSheetName

   'Set the orientation and scaling
   Set m_objPageSetup = m_objWorksheet.PageSetup
   With m_objPageSetup
      Select Case PageOrientation
         Case xlPortrait:  .Orientation = xlPortrait
         Case Else:        .Orientation = xlLandscape
      End Select
      m_lxPageOrientation = .Orientation
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = 100  'Set this high so that it will print 1 page wide
   End With

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* OpenExcelFileWithNoWorksheets
'* - Returns a fully qualified Excel file name based on the export_path in the
'*   INI file.  Deletes existing file if applicable.
'*--------------------------------------------------------------

Public Sub OpenExcelFileWithNoWorksheets(FileName As String, _
    Optional PageOrientation As XlPageOrientation = xlLandscape)

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "OpenExcelFileWithNoWorksheets"
   Dim sReturnString     As String
   Dim lRc               As Long
   Dim i As Integer

   m_strExcelFileName = FileName

   m_lxPageOrientation = PageOrientation

   If Right$(m_strExcelFileName, 4) <> ".xls" Then
      m_strExcelFileName = m_strExcelFileName & ".xls"
   End If

   If Exists(m_strExcelFileName) Then
      Kill m_strExcelFileName
   End If

   Set m_objXL = CreateObject("Excel.Application")
   Set m_objWorkbook = m_objXL.Workbooks.Add

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* AddNewWorkSheet
'* - Creates a new worksheet
'*--------------------------------------------------------------

Public Sub AddNewWorksheet(WorkSheetName As String)

   Set m_objWorksheet = m_objXL.Worksheets.Add
   m_objWorksheet.Name = WorkSheetName
   m_objWorksheet.Activate

   'Set the orientation and scaling
   Set m_objPageSetup = m_objWorksheet.PageSetup
   With m_objPageSetup
      Select Case m_lxPageOrientation
         Case xlPortrait:  .Orientation = xlPortrait
         Case Else:        .Orientation = xlLandscape
      End Select
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = 100  'Set this high so that it will print 1 page wide
   End With

   m_lngExcelRow = 0

End Sub

'*--------------------------------------------------------------
'* CloseExcelFile
'* - Closes the Excel file that was previously opened and clears memory.
'*--------------------------------------------------------------

Public Sub CloseExcelFile()

   On Error GoTo ErrorHandler
   Const constErrRoutine As String = "CloseExcelFile"

   If m_objWorksheet Is Nothing Then
      Err.Raise 1000, "ttExcel.CreateExcelFromADORecordset", _
          "You must first invoke the OpenExcelFile method before calling this " _
          & "method."
      GoTo ExitMe
   End If

   Call StoreExcelRow(A:=" ")

   m_objWorkbook.SaveAs m_strExcelFileName
   m_objWorkbook.Close SaveChanges:=False
   m_objXL.Quit
   Set m_objWorksheet = Nothing
   Set m_objWorkbook = Nothing
   Set m_objXL = Nothing

ExitMe:
   Exit Sub

ErrorHandler:
   Err.Raise Err.Number, m_constPgm & constErrRoutine, Err.Description
   Resume ExitMe

End Sub

'*--------------------------------------------------------------
'* Exists
'* - Determines if a file exists. Returns true or false
'* - Example      : blnDoesItExist = Exists("C:\CONFIG.SYS")
'*--------------------------------------------------------------

Private Function Exists(FileName) As Integer

  Dim lFileLen    As Long

  On Error Resume Next
  lFileLen = FileLen(FileName)

  If lFileLen > 0 Then
    Exists = True
  End If

End Function

'*--------------------------------------------------------------
'* Class_Terminate
'* -  Close the excel file and cleanup if they forget to.
'*--------------------------------------------------------------
Private Sub Class_Terminate()

   If m_objXL Is Nothing Then
      'Cool, get out
   Else
      'They forgot to close the file
      CloseExcelFile
   End If

End Sub

'*--------------------------------------------------------------
'* ReadValue
'* -  Reads a value from the Excel worksheet based on a previously defined label
'*--------------------------------------------------------------

Public Sub ReadValue(Label As String, ByRef Value As String)

   On Error Resume Next

   Value = m_objXL.Range(Label).Value

   If Err.Number = 0 Then
      ' It worked, just return
   Else
      If Err.Number = m_lngExcelLabelNotFound Then
         Value = ""
      Else
         Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, _
             Err.HelpContext
      End If
   End If

End Sub

'###########################################################
'#
'#   This class has been brought to you by Pragmatic Software Co. Inc,
'#   the creators of Defect Tracker, the tool of choice for tracking
'#   functional specifications, test cases and software bugs.
'#   Learn more at http://www.DefectTracker.com.
'#   Affiliate program also available at
'#          http://www.PragmaticSW.com/AffiliateSignup.
'#
'###########################################################


