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
Public Function extractWords(ByRef sp_error As String, sp_filepath As String) As Integer
On Error GoTo Hell
  '----------------------------------------------------------------
  Dim pdf As New PDFPARSERLib.Document
  Dim content As PDFPARSERLib.content
  Dim text As PDFPARSERLib.text
  Dim X As Single, Y As Single, FontSize As Single
  '----------------------------------------------------------------
  sp_error = ""
  '----------------------------------------------------------------
  If pdf.Open(sp_filepath) Then
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
Private Function findArticleTitle(ip_start_index As Integer, _
                                  ByRef ip_article_start_index As Integer, _
                                  ByRef ip_article_end_index) As String
  Dim i As Integer
  Dim maxSize As Single
  Dim maxSize_index As Integer 'index of the PDFWord in the list that has maximum size
  maxSize = 0
  If ip_start_index < 1 Then ip_start_index = 1
  For i = ip_start_index To m_pdfWords.Count
    Dim t_pdfWord As pdfWord
    Set t_pdfWord = m_pdfWords.Item(i)
    If t_pdfWord.m_FontSize > maxSize Then
      maxSize = t_pdfWord.m_FontSize
      maxSize_index = i
    End If
  Next
  'now we have index of the first word of Article Title and its size
  'so get the complete title of the Article having size same as max size
  Dim sArticleTitle As String
  ip_article_start_index = maxSize_index
  If (maxSize_index > 0) Then
    For i = maxSize_index To m_pdfWords.Count
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
'function to find the YEAR
'returns the index of the year word from collection
Private Function findYear() As Integer
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
  findYear = iYearIndex
End Function
'#####################################################################################
'function to find the Page Number
'returns the index of the Page number word from collection
Private Function findPageNumber() As Integer
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
      If IsNumeric(sStartPageNum) And IsNumeric(sEndPageNum) Then
        'page number found
        iPageNumIndex = i
        Exit For
      End If
    End If
  Next i
  findPageNumber = iPageNumIndex
End Function
'#####################################################################################
'the parameter passed is equal to (index of the last word from article title) + 1
Private Function findFirstAuthor(ip_start_index As Integer) As String
  If ip_start_index < 1 Or ip_start_index > m_pdfWords.Count Then
    findFirstAuthor = ""
    Exit Function
  End If
  Dim sAuthors As String
  'first find all authors and then extract first author
  sAuthors = getAuthors(ip_start_index)
  'MsgBox "All Authors ---> " & sAuthors
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
'this function removes all special characters except comma (,) from the authors string
'#####################################################################################
'the parameter passed is equal to (index of the last word from article title) + 1
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
'    If GeneralMod.isSpecialChar(Trim(t_pdfWord.m_word)) Then
'      sAuthors = sAuthors & t_pdfWord.m_word & " "
'    Else
'      sAuthors = sAuthors & t_pdfWord.m_word & " "
'    End If
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

'parse the contents of the PDF file
'returns an object of PDFFileDetails class with all details of PDF file
Public Function parsePDFContents(ByRef sp_error As String, sp_filepath As String) As PDFFileDetails
On Error GoTo PDFParse_Error
  Dim sMessage As String
  Dim t_pdfWord As pdfWord
  '----------------------------------------------------------------
  sp_error = ""
  '----------------------------------------------------------------
  'finding year
  Dim iYearIndex As Integer
  Dim sYear As String
  iYearIndex = findYear
  If iYearIndex <> -1 Then
    'MsgBox ("Year ----> " & m_pdfWords.Item(iYearIndex).m_word)
    sYear = m_pdfWords.Item(iYearIndex).m_word
    sMessage = sMessage & "Year ----> " & sYear & vbNewLine
  Else
    'MsgBox "Year not found."
    sMessage = sMessage & "Year not found." & vbNewLine
  End If
  '----------------------------------------------------------------
  'finding article title
  Dim sArticle_title As String
  Dim article_start_index As Integer
  Dim article_end_index As Integer
  'MsgBox ("Title of the Article ---> " & PDFParserMod.findArticleTitle(iYearIndex, article_start_index, article_end_index))
  sArticle_title = PDFParserMod.findArticleTitle(iYearIndex, article_start_index, article_end_index)
  sMessage = sMessage & "Title of the Article ---> " & sArticle_title & vbNewLine
  '----------------------------------------------------------------
  'finding first author
  Dim sFirst_author As String
  'MsgBox "First Author ----> " & findFirstAuthor(article_end_index + 1)
  sFirst_author = findFirstAuthor(article_end_index + 1)
  sMessage = sMessage & "First Author ----> " & sFirst_author & vbNewLine
  '----------------------------------------------------------------
  'finding page number
  Dim sPage_no As String
  Dim page_num_index As Integer
  page_num_index = findPageNumber()
  If page_num_index <> -1 Then
    'MsgBox "Page number ----> " & m_pdfWords.Item(page_num_index).m_word
    sPage_no = m_pdfWords.Item(page_num_index).m_word
    sMessage = sMessage & "Page number ----> " & sPage_no & vbNewLine
  Else
    'MsgBox "Page Number not found."
    sMessage = sMessage & "Page Number not found." & vbNewLine
  End If
  '----------------------------------------------------------------
  'finding volume number
  Dim iVolNumIndex As Integer
  Dim sVolNum As String
  sVolNum = ""
  For Index = iYearIndex - 1 To 1 Step -1
    Set t_pdfWord = m_pdfWords.Item(Index)
    If IsNumeric(t_pdfWord.m_word) Then
      iVolNumIndex = Index
      'MsgBox "Volume Number ---> " & t_pdfWord.m_word
      sVolNum = t_pdfWord.m_word
      Exit For
    End If
  Next
  If sVolNum <> "" Then
    sMessage = sMessage & "Volume Number ---> " & t_pdfWord.m_word & vbNewLine
  Else
    sMessage = sMessage & "Volume Number not found." & vbNewLine
  End If
  '----------------------------------------------------------------
  'finding journal title
  Dim sJournal_title As String
  'MsgBox ("Title of the Journal ---> " & PDFParserMod.findJournalTitle(iVolNumIndex - 1))
  sJournal_title = PDFParserMod.findJournalTitle(iVolNumIndex - 1)
  sMessage = sMessage & "Title of the Journal ---> " & sJournal_title & vbNewLine
  '----------------------------------------------------------------
  'separate out file name, file path and journal subject
  Dim sFilename As String
  Dim sFilePath As String
  Dim sJournal_subject As String
  
  Dim sTemp As String
  Dim iIndex As Integer
  sTemp = StrReverse(sp_filepath)
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
  'frmLoadPDF.logMessage (sMessage)
  Exit Function
  '----------------------------------------------------------------
PDFParse_Error:
  sp_error = ErrorHandlingMod.createErrorMsg("Error while parsing the file : " & sp_filepath, _
                                             Err.Number, Err.Description)
  Set parsePDFContents = Nothing
End Function
'#####################################################################################

