Attribute VB_Name = "SearchMod"
'#################################################################################
Public Function getSearchQuery(bp_IsAndOpr As Boolean, _
                                sp_ArticleTitle As String, sp_Filename As String, _
                                sp_FirstAuthor As String, sp_JournalSubject As String, _
                                sp_JournalTitle As String, sp_PageNum As String, _
                                sp_VolNum As String, sp_Year As String, _
                                sp_Notes As String, sp_Filepath As String) As String
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
  If (bp_IsAndOpr = True) Then
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
  If (Len(Trim(sp_ArticleTitle)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.article_title) like """ & _
                  UCase$((Replace(Trim(sp_ArticleTitle), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Filename
  If (Len(Trim(sp_Filename)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.filename like """ & _
                  UCase$(Replace(Trim(sp_Filename), "*", "%")) & "%"""
  End If
  '------------------------------------------
  'Search on First Author
  If (Len(Trim(sp_FirstAuthor)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.first_author like """ & _
                  UCase$((Replace(Trim(sp_FirstAuthor), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Journal Subject
  If (Len(Trim(sp_JournalSubject)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.journal_subject like """ & _
                  UCase$((Replace(Trim(sp_JournalSubject), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Journal Title
  If (Len(Trim(sp_JournalTitle)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.journal_title like """ & _
                  UCase$((Replace(Trim(sp_JournalTitle), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Page Number
  If (Len(Trim(sp_PageNum)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.page_no like """ & _
                  UCase$((Replace(Trim(sp_PageNum), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Volume Number
  If (Len(Trim(sp_VolNum)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.volume_no like """ & _
                  UCase$((Replace(Trim(sp_VolNum), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Year
  If (Len(Trim(sp_Year)) > 0) Then
    sSQL = sSQL & sLogicalOp & "PDFFileDetails.[year] like """ & _
                  UCase$((Replace(Trim(sp_Year), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on Notes
  If (Len(Trim(sp_Notes)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.notes) like """ & _
                  UCase$((Replace(Trim(sp_Notes), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  'Search on File path
  If (Len(Trim(sp_Filepath)) > 0) Then
    sSQL = sSQL & sLogicalOp & "ucase(PDFFileDetails.filepath) like """ & _
                  UCase$((Replace(Trim(sp_Filepath), "*", "%"))) & "%"""
  End If
  '------------------------------------------
  sSQL = sSQL & " ORDER BY PDFFileDetails.article_title ASC;"
  getSearchQuery = sSQL
End Function
'#################################################################################

