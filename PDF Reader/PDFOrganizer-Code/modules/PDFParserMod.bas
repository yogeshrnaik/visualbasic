Attribute VB_Name = "PDFParserMod"
'#####################################################################################
'This module is used to parse the contents of the PDF file and
'extract the required field from its contents
'e.g. Journal Title, Article Title, Volume No., Year, etc.
'#####################################################################################
'variable used to store all words extracted from the PDF file
Private m_pdfWords As Collection
'#####################################################################################
'function to extract words from a pdf file
'all these eextracted words are stored in the collection object "m_pdfWords"
Public Function extractWords(ByRef sp_error As String, sp_Filepath As String) As Integer
On Error GoTo Hell
  '----------------------------------------------------------------
  Dim pdf As New PDFPARSERLib.Document
  Dim content As PDFPARSERLib.content
  Dim text As PDFPARSERLib.text
  Dim X As Single, Y As Single, FontSize As Single
  '----------------------------------------------------------------
  sp_error = ""
  '----------------------------------------------------------------
  If pdf.Open(sp_Filepath) Then
    pdf.PageNo = 1                    ' set the current page number
    Set content = pdf.Page.content          ' get the page's content
    If Not (content Is Nothing) Then
      content.BreakWords = True           ' extract words
      content.Reset True
      'Me.txtPageContents.text = Me.txtPageContents.text & "- - - Page " & CurPage & " - - -" & Chr(13) & Chr(10)
      If m_pdfWords Is Nothing Then
        Set m_pdfWords = New Collection
      ElseIf m_pdfWords.Count > 0 Then
        Dim i As Integer
        While m_pdfWords.Count > 0
          m_pdfWords.Remove (1)
        Wend
      End If
      Do
        If content.GetNextText Is Nothing Then Exit Do
        Set text = content.text                           'at this point you can access the text properties
        'no need to get the text that is after Abstract
        'frmLoadPDF.txtLog.text = frmLoadPDF.txtLog.text & vbNewLine & text.RawString
        If UCase(text.UnicodeString) = UCase("Abstract") Then Exit Do
        m_pdfWords.Add PDFParserMod.getPDFWord(text)        'add the extracted text in Collection
'       Me.txtPageContents.text = Me.txtPageContents.text & text.UnicodeString & _
'                                 " " & " :  " & X & "," & Y _
'                                 & "  Size: " & FontSize & Chr(13) & Chr(10)
      Loop
    Else
      'MsgBox "There is no content on this page", vbExclamation, GeneralMod.getApplnName()
    End If
    'frmCaptureText.txtPageContents.text = frmCaptureText.txtPageContents.text & vbNewLine
    pdf.Close
    extractWords = GeneralMod.OKAY
  Else
    'MsgBox "Couldn't open input file", vbExclamation, GeneralMod.getApplnName()
    extractWords = GeneralMod.ERROR_OCCURED
  End If
  Exit Function
Hell:
'  MsgBox "Error has occured while extracting words from PDF file." & vbNewLine & _
'         "Error Number: " & Err.Number & vbNewLine & _
'         "Error Description: " & Err.Description, vbCritical, GeneralMod.getApplnName
  sp_error = "Error has occured while extracting words from PDF file." & vbNewLine & _
             "Error Number: " & Err.Number & vbNewLine & _
             "Error Description: " & Err.Description
  extractWords = GeneralMod.ERROR_OCCURED
End Function
'#####################################################################################
Public Sub displayExtractedWords()
  Dim i As Integer
  Dim pdfWord As pdfWord
  If m_pdfWords Is Nothing Then
    'MsgBox "No words extracted.", vbExclamation, GeneralMod.getApplnName
    Exit Sub
  ElseIf m_pdfWords.Count = 0 Then
    'MsgBox "No words extracted.", vbExclamation, GeneralMod.getApplnName
    Exit Sub
  End If
  'frmCaptureText.txtPageContents.text = ""
  For i = 1 To m_pdfWords.Count
    Set pdfWord = m_pdfWords.Item(i)
'    frmCaptureText.txtPageContents.text = frmCaptureText.txtPageContents.text & pdfWord.m_word & _
'                                          " " & " :  " & pdfWord.m_X & "," & pdfWord.m_Y & _
'                                          "  Size: " & pdfWord.m_FontSize & Chr(13) & Chr(10)
  Next i
End Sub
'#####################################################################################
Public Function getPDFWord(p_pdfText As PDFPARSERLib.text) As pdfWord
  Dim pw As New pdfWord
  pw.init Trim(p_pdfText.UnicodeString), Round(p_pdfText.TextMatrix.e, 1), _
          Round(p_pdfText.TextMatrix.f, 1), Round(p_pdfText.TextMatrix.d, 1)
  
  Set getPDFWord = pw
'  m_FontSize = Round(text.TextMatrix.d, 1) ' the font size
'  m_X = Round(text.TextMatrix.e, 1) ' the X position
'  m_Y = Round(text.TextMatrix.f, 1) ' the y position
'  m_word = text.UnicodeString
End Function
'#####################################################################################
'function to decide the title of the article
'Logic: the text in largest font is the Article Title
Private Function findArticleTitle(ByRef ip_article_end_index) As String
  Dim i As Integer
  Dim maxSize As Single
  Dim article_start_index As Integer
  Dim maxSize_index As Integer 'index of the PDFWord in the list that has maximum size
  maxSize = 0
  For i = 1 To 100
    If (i < m_pdfWords.Count) Then
      Dim t_pdfWord As pdfWord
      Set t_pdfWord = m_pdfWords.Item(i)
      If t_pdfWord.m_FontSize > maxSize Then
        maxSize = t_pdfWord.m_FontSize
        maxSize_index = i
      End If
    End If
  Next
  'now we have index of the first word of Article Title and its size
  'so get the complete title of the Article having size same as max size
  Dim sArticleTitle As String
  article_start_index = maxSize_index
  If (article_start_index > 0) Then
    For i = article_start_index To m_pdfWords.Count
      Set t_pdfWord = m_pdfWords.Item(i)
      If t_pdfWord.m_FontSize <> maxSize Then
        ip_article_end_index = i - 1
        Exit For
      Else
        sArticleTitle = sArticleTitle & t_pdfWord.m_word & " "
      End If
    Next i
  End If
  findArticleTitle = sArticleTitle
End Function
'#####################################################################################
'function to find the YEAR which is in parenthesis e.g. (2005)
'returns the index of the year word from collection
Private Function findYearByBracket() As Integer
  Dim i As Integer
  Dim iYearIndex As Integer
  Dim t_pdfWord As pdfWord
  iYearIndex = -1
  For i = 1 To m_pdfWords.Count
    Set t_pdfWord = m_pdfWords.Item(i)
    If Len(t_pdfWord.m_word) = 6 Then
      Dim first_char As String
      Dim last_char As String
      first_char = Mid(t_pdfWord.m_word, 1, 1)
      last_char = Mid(StrReverse(t_pdfWord.m_word), 1, 1)
      If first_char = "(" And last_char = ")" Then
        'year found
        iYearIndex = i
        Exit For
      End If
    End If
  Next i
  findYearByBracket = iYearIndex
End Function
'#####################################################################################
'function to find the Page Number
'Logic: Find page number by identifying hypen
'returns the index of the Page number word from collection
'also sets the boolean variable "includesYearVol" to true if Page Num
'has got Year and Vol. no. as well
'E.g. in case of "Year;VolNo:StartPgNo-EndPgNo" as in "2004;17:200–206"
Private Function findPageNumber(ByRef sPage_no As String, _
                                ByRef includesYearVol As Boolean) As Integer
  Dim i As Integer
  Dim iHyphenIndex As Integer
  Dim iPageNumIndex As Integer
  Dim sStartPageNum As String
  Dim sEndPageNum As String
  Dim t_pdfWord As pdfWord
  
  iPageNumIndex = -1
  For i = 1 To m_pdfWords.Count
    Set t_pdfWord = m_pdfWords.Item(i)
    'find the index of Hyphen (-)
    iHyphenIndex = InStr(1, t_pdfWord.m_word, "–", vbTextCompare)
    If iHyphenIndex <> 0 Then
      'separate the word based on iHyphenIndex
      sStartPageNum = Mid(t_pdfWord.m_word, 1, iHyphenIndex - 1)
      sEndPageNum = Mid(t_pdfWord.m_word, iHyphenIndex + 1)
      If IsNumeric(sEndPageNum) Then
        If IsNumeric(sStartPageNum) Then
          'page number found
          sPage_no = t_pdfWord.m_word
          iPageNumIndex = i
          includesYearVol = False
          Exit For
        Else
          'may be the sStartPageNum has got Year and Vol. no. as well
          'e.g. in case of "Year;VolNo:StartPgNo-EndPgNo" as in "2004;17:200–206"
          'so, traverse back words in the sStartPageNum and extract all numbers
          Dim j As Integer
          Dim last_char As String
          Dim sFinalStartPgNo As String
          j = 0
          Do
            last_char = Mid(StrReverse(sStartPageNum), 1, 1)
            If IsNumeric(last_char) Then
              sFinalStartPgNo = last_char + sFinalStartPgNo
              sStartPageNum = Mid(sStartPageNum, 1, Len(sStartPageNum) - 1)
            Else
              Exit Do
            End If
            j = j + 1
          Loop
          If j > 0 Then
            'some no was found
            iPageNumIndex = i
            includesYearVol = True
            sPage_no = sFinalStartPgNo + "-" + sEndPageNum
            Exit For
          End If
        End If
      End If
    End If
  Next i
  findPageNumber = iPageNumIndex
End Function
'#####################################################################################
'the parameter passed is equal to (index of the last word from article title) + 1
'Logic: Find First Author as the text after Article title till font change
Private Function findFirstAuthor(ip_start_index As Integer) As String
  If ip_start_index < 1 Or ip_start_index > m_pdfWords.Count Then
    findFirstAuthor = ""
    Exit Function
  End If
  Dim sAuthors As String
  'first find all authors and then extract first author
  sAuthors = getAuthors(ip_start_index)
  If sAuthors = "" Then
    findFirstAuthor = ""
    Exit Function
  End If
  'extract first author from all authors string now
  Dim iCommaIndex As Integer
  iCommaIndex = InStr(1, sAuthors, ",", vbTextCompare)
  If iCommaIndex <> 0 Then
    sAuthors = Mid(sAuthors, 1, iCommaIndex - 1)
  End If
  findFirstAuthor = sAuthors
End Function
'#####################################################################################
'Logic: Find all authors as the text after Article title till font change
'the parameter passed is equal to (index of the last word from article title) + 1
'this function removes all special characters except comma (,) from the authors string
Private Function getAuthors(ip_start_index As Integer) As String
  If ip_start_index < 1 Or ip_start_index > m_pdfWords.Count Then
    getAuthors = ""
    Exit Function
  End If
  Dim sAuthors As String
  Dim i As Integer
  Dim t_size As Single
  Dim t_pdfWord As pdfWord
  t_size = 0
  For i = ip_start_index To m_pdfWords.Count
    Set t_pdfWord = m_pdfWords.Item(i)
    If t_size = 0 Then
      t_size = t_pdfWord.m_FontSize
    ElseIf t_size <> t_pdfWord.m_FontSize And Len(Trim(t_pdfWord.m_word)) > 1 Then
      'the condition Len(Trim(t_pdfWord.m_word)) > 1 is because '
      'sometimes after one author's name there is a superscript of one character
      'this condition can be removed if user wants only first author
      Exit For
    End If
    sAuthors = sAuthors & t_pdfWord.m_word & " "
  Next
  sAuthors = GeneralMod.removeSpecialChars(sAuthors, "' . - ,")
  sAuthors = Replace(sAuthors, "  ", "")
  getAuthors = sAuthors
End Function
'#####################################################################################
'function to decide the title of the journal
Private Function findJournalTitle(ip_start_index As Integer) As String
  If ip_start_index < 1 Or ip_start_index > m_pdfWords.Count Then
    findJournalTitle = ""
    Exit Function
  End If

  Dim i As Integer
  Dim t_Y As Single
  Dim t_pdfWord As pdfWord
  t_Y = -1
  For i = ip_start_index To 1 Step -1
    Set t_pdfWord = m_pdfWords.Item(i)
    If t_Y = -1 Then
      t_Y = t_pdfWord.m_Y
    ElseIf t_Y <> t_pdfWord.m_Y Then
      Exit For
    End If
    sJournalTitle = t_pdfWord.m_word & " " & sJournalTitle
  Next
  findJournalTitle = sJournalTitle
End Function
'#####################################################################################
'In case of format "JournalTitle Year:Vol:PageNum"
'the Year and Volume number are embedded in word containing Page Number
'e.g. Ophthalmology 2003;110:765–771
'so this method finds the year and volume number using page number index
Private Sub findYearVolByPgNumIndex(ByVal ip_page_num_index As Integer, _
                                    ByRef sp_Year As String, ByRef sp_vol_num As String)
  If (ip_page_num_index < 1) Then
    sp_Year = ""
    sp_vol_num = ""
    Exit Sub
  End If
  Dim sPageNumWord As String
  Dim t_pdfWord As pdfWord
  Set t_pdfWord = m_pdfWords.Item(ip_page_num_index)
  sPageNumWord = t_pdfWord.m_word
  
  Dim iSemiColonIndex As Integer
  Dim iColonIndex As Integer
  iSemiColonIndex = InStr(1, sPageNumWord, ";")
  iColonIndex = InStr(1, sPageNumWord, ":")
  If (iSemiColonIndex < 1 Or iColonIndex < 1) Then
    sp_Year = ""
    sp_vol_num = ""
    Exit Sub
  End If
  sp_vol_num = Mid(sPageNumWord, iSemiColonIndex + 1, iColonIndex - iSemiColonIndex - 1)
  If Not IsNumeric(sp_vol_num) Then
    sp_vol_num = ""
  End If
  'to find year, traverse from iSemiColonIndex till some text other than number is found
  'e.g. Ophthalmology 2003;110:765–771
  Dim i As Integer
  Dim one_char As String
  For i = iSemiColonIndex - 1 To 1 Step -1
    one_char = Mid(sPageNumWord, i, 1)
    If (IsNumeric(one_char)) Then
      sp_Year = one_char & sp_Year
    End If
  Next
  If Not IsNumeric(sp_Year) Then
    sp_Year = ""
  End If
End Sub
'#####################################################################################
'parse the contents of the PDF file
'returns an object of PDFFileDetails class with all details of PDF file
'Logic:
'(1) First find the Article Title as the largest text in PDF
'(2) Find First Author as the text after Article title till font change
'(3) Find page number by identifying hypen
'(4) If Year and Volume number are embedded in word containing Page Number
'     then it is likely that the format is "JournalTitle Year:Vol:PageNum"
'     e.g. Ophthalmology 2003;110:765–771
'     so use index of page number, to find year and volume no.
'(5) If Year is not found by page number index, find Year by bracket () e.g. (2004)
'(6) If volume number is not found yet, and if year has been found by brackets, then
'     if is likely that the format is "JournalTitle Vol (Year) PageNum"
'     E.g. Contact Lens & Anterior Eye 27 (2004) 209–212
'     Hence, find volume number using year index
'(7) Find Journal Title depending on Year index or Volume number index or Page number index
'     if YearIndex < Volume Index < PageNumIndex
'     Use year, volume number and page number index in asending order
Public Function parsePDFContents(ByRef sp_error As String, sp_Filepath As String) As PDFFileDetails
On Error GoTo PDFParse_Error
  Dim sMessage As String
  Dim t_pdfWord As pdfWord
  '----------------------------------------------------------------
  sp_error = ""
  '----------------------------------------------------------------
  '(1) First find the Article Title as the largest text in PDF
  Dim sArticle_title As String
  Dim article_end_index As Integer
  sArticle_title = findArticleTitle(article_end_index)
  sMessage = sMessage & "Title of the Article ---> " & sArticle_title & vbNewLine
  '----------------------------------------------------------------
  '(2) Find First Author as the text after Article title till font change
  Dim sFirst_author As String
  sFirst_author = findFirstAuthor(article_end_index + 1)
  sMessage = sMessage & "First Author ----> " & sFirst_author & vbNewLine
  '----------------------------------------------------------------
  '(3) Find page number by identifying hypen
  Dim sPage_no As String
  Dim iPage_num_index As Integer
  Dim iYearIndex As Integer
  Dim sYear As String
  Dim iVolNumIndex As Integer
  Dim sVolNum As String
  Dim includesYearVol As Boolean
  iYearIndex = -1
  iVolNumIndex = -1
  sYear = ""
  sVolNum = ""
  sPage_no = ""
  iPage_num_index = findPageNumber(sPage_no, includesYearVol)
  If Len(sPage_no) = 0 Then
    iPage_num_index = -1
    sMessage = sMessage & "Page Number not found." & vbNewLine
  Else
    sMessage = sMessage & "Page number ----> " & sPage_no & vbNewLine
    If (includesYearVol) Then
      'find Year and Volume number by Page Number Index
      findYearVolByPgNumIndex iPage_num_index, sYear, sVolNum
    End If
  End If
  '----------------------------------------------------------------
  'finding year only when Year is not yet found
  If Len(sYear) = 0 Then
    iYearIndex = findYearByBracket
    If iYearIndex <> -1 Then
      sYear = m_pdfWords.Item(iYearIndex).m_word
      sMessage = sMessage & "Year ----> " & sYear & vbNewLine
    Else
      sMessage = sMessage & "Year not found." & vbNewLine
    End If
  End If
  '----------------------------------------------------------------
  'finding volume number only when volume number is not yet found
  If Len(sVolNum) = 0 Then
    sVolNum = ""
    For Index = iYearIndex - 1 To 1 Step -1
      Set t_pdfWord = m_pdfWords.Item(Index)
      If IsNumeric(t_pdfWord.m_word) Then
        iVolNumIndex = Index
        sVolNum = t_pdfWord.m_word
        Exit For
      End If
    Next
    If sVolNum <> "" Then
      sMessage = sMessage & "Volume Number ---> " & t_pdfWord.m_word & vbNewLine
    Else
      sMessage = sMessage & "Volume Number not found." & vbNewLine
    End If
  End If
  '----------------------------------------------------------------
  'finding journal title
  Dim sJournal_title As String
  'MsgBox ("Title of the Journal ---> " & PDFParserMod.findJournalTitle(iVolNumIndex - 1))
  
  If iVolNumIndex <> -1 Then
    sJournal_title = PDFParserMod.findJournalTitle(iVolNumIndex - 1)
  End If
  If Len(sJournal_title) = 0 And iYearIndex <> -1 Then
    sJournal_title = PDFParserMod.findJournalTitle(iYearIndex - 1)
  End If
  If Len(sJournal_title) = 0 And iPage_num_index <> -1 Then
    sJournal_title = PDFParserMod.findJournalTitle(iPage_num_index - 1)
  End If
  sMessage = sMessage & "Title of the Journal ---> " & sJournal_title & vbNewLine
  '----------------------------------------------------------------
  'separate out file name, file path and journal subject
  Dim sFilename As String
  Dim sFilePath As String
  Dim sJournal_subject As String
  
  Dim sTemp As String
  Dim iIndex As Integer
  sTemp = StrReverse(sp_Filepath)
  iIndex = InStr(1, sTemp, "\", vbTextCompare)
  sFilename = StrReverse(Mid(sTemp, 1, iIndex - 1))
  sFilePath = StrReverse(Mid(sTemp, iIndex))
  
  'use length - 1 because, file path has "\" at the end
  sTemp = StrReverse(Mid(sFilePath, 1, Len(sFilePath) - 1))
  iIndex = InStr(1, sTemp, "\", vbTextCompare)
  If (iIndex = 0) Then
    sJournal_subject = StrReverse(sTemp)
  Else
    sJournal_subject = StrReverse(Mid(sTemp, 1, iIndex - 1))
  End If
'  MsgBox (sTemp & vbNewLine & iIndex & vbNewLine & sFilename & vbNewLine & vbNewLine & _
'          sFilePath & vbNewLine & vbNewLine & sJournal_subject)
  '----------------------------------------------------------------
  Dim oPDFDetails As New PDFFileDetails
  oPDFDetails.init sFilename, sFilePath, sArticle_title, sJournal_title, sJournal_subject, sYear, _
                   sVolNum, sPage_no, sFirst_author, ""
  Set parsePDFContents = oPDFDetails
  'MsgBox sMessage
  frmLoadPDF.logMessage (sMessage)
  Exit Function
  '----------------------------------------------------------------
PDFParse_Error:
  sp_error = ErrorHandlingMod.createErrorMsg("Error while parsing the file : " & sp_Filepath, _
                                             Err.Number, Err.Description)
  Set parsePDFContents = Nothing
End Function
'#####################################################################################

