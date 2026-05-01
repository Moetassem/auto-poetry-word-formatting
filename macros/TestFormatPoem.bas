Attribute VB_Name = "TestFormatPoem"
' =====================================================================
' TestFormatPoem - Unit and integration tests for FormatPoem module
' =====================================================================
' HOW TO RUN
'   1. In Word, open the VBA editor (Alt+F11).
'   2. Import this file (File > Import File...).
'   3. Make sure FormatPoem.bas is also imported.
'   4. Press F5 or run  RunAllTests  from the Macros dialog (Alt+F8).
'   5. Results appear in the Immediate Window (Ctrl+G) and a summary
'      MsgBox at the end.
'
' IMPORTANT: These tests create and immediately destroy temporary Word
' documents. Do not run them with unsaved work open if Word is set to
' ask about saving on close for every document.
'
' NOTE: Private functions in FormatPoem (StripTrailingCR, GetMode,
' SetMode, ConvertLineToPoetryTable, FillCell, BindEnter, UnbindEnter)
' cannot be called directly from another module. They are tested either
' (a) via local reimplementations that mirror the spec, or (b) through
' the observable effects of the public API.
' =====================================================================

Option Explicit

' ---- constants duplicated from FormatPoem for use in assertions ----
Private Const STATE_VAR_NAME As String = "ArabicPoetryTableMode"
Private Const POETRY_SEPARATOR As String = "**"

' ---- test counters ----
Private mPassed As Long
Private mFailed As Long

' =====================================================================
' Entry point
' =====================================================================
Public Sub RunAllTests()
    mPassed = 0
    mFailed = 0

    Debug.Print String(60, "=")
    Debug.Print "TestFormatPoem  " & Now()
    Debug.Print String(60, "=")

    ' Pure-logic suites (no Word document required)
    Suite_StripTrailingCR
    Suite_SeparatorParsing

    ' Integration suites (create/destroy temporary documents)
    Suite_StateManagement
    Suite_FormatOnEnter_NoSeparator
    Suite_FormatOnEnter_WithSeparator
    Suite_FormatOnEnter_InTable
    Suite_TableStructure
    Suite_EdgeCases

    Debug.Print String(60, "-")
    Debug.Print "TOTAL  passed=" & mPassed & "  failed=" & mFailed
    Debug.Print String(60, "=")

    MsgBox "TestFormatPoem results" & vbCrLf & vbCrLf & _
           "Passed : " & mPassed & vbCrLf & _
           "Failed : " & mFailed & vbCrLf & vbCrLf & _
           "See Immediate Window (Ctrl+G) for details.", _
           IIf(mFailed = 0, vbInformation, vbCritical), _
           "TestFormatPoem"
End Sub

' =====================================================================
' Assert helpers
' =====================================================================
Private Sub Assert(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        mPassed = mPassed + 1
        Debug.Print "  [PASS]  " & testName
    Else
        mFailed = mFailed + 1
        Debug.Print "  [FAIL]  " & testName
    End If
End Sub

Private Sub AssertEqual(ByVal actual As String, ByVal expected As String, _
                        ByVal testName As String)
    If actual = expected Then
        mPassed = mPassed + 1
        Debug.Print "  [PASS]  " & testName
    Else
        mFailed = mFailed + 1
        Debug.Print "  [FAIL]  " & testName & _
                    "  (expected=" & Chr(34) & expected & Chr(34) & _
                    "  got=" & Chr(34) & actual & Chr(34) & ")"
    End If
End Sub

Private Sub AssertLong(ByVal actual As Long, ByVal expected As Long, _
                       ByVal testName As String)
    If actual = expected Then
        mPassed = mPassed + 1
        Debug.Print "  [PASS]  " & testName
    Else
        mFailed = mFailed + 1
        Debug.Print "  [FAIL]  " & testName & _
                    "  (expected=" & expected & "  got=" & actual & ")"
    End If
End Sub

' =====================================================================
' Local reimplementation of StripTrailingCR
' =====================================================================
' StripTrailingCR is Private in FormatPoem, so it cannot be called
' from here. The function below mirrors the published specification
' (FormatPoem.bas lines 217-225) and is used to verify the algorithm
' independently of Word's object model.
Private Function LocalStripCR(ByVal s As String) As String
    If Len(s) > 0 Then
        If Right$(s, 1) = Chr(13) Then
            LocalStripCR = Left$(s, Len(s) - 1)
            Exit Function
        End If
    End If
    LocalStripCR = s
End Function

' =====================================================================
' Local reimplementation of the sadr/ajuz split logic
' =====================================================================
' ConvertLineToPoetryTable is Private. This mirrors lines 101-103 of
' FormatPoem.bas to let us unit-test the split algorithm in isolation.
Private Sub LocalSplit(ByVal lineText As String, ByVal sepPos As Long, _
                       ByRef outSadr As String, ByRef outAjuz As String)
    Dim sepLen As Long
    sepLen = Len(POETRY_SEPARATOR)
    outSadr = Trim$(Left$(lineText, sepPos - 1))
    outAjuz = Trim$(Mid$(lineText, sepPos + sepLen))
End Sub

' =====================================================================
' Helper: create a temporary blank document, return it
' =====================================================================
Private Function NewTempDoc() As Document
    Set NewTempDoc = Documents.Add(Visible:=False)
End Function

' =====================================================================
' Helper: close a temp doc without saving
' =====================================================================
Private Sub CloseTempDoc(ByVal doc As Document)
    On Error Resume Next
    doc.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub

' =====================================================================
' Helper: read the state variable from a document
' =====================================================================
Private Function ReadStateVar(ByVal doc As Document) As String
    Dim s As String
    s = ""
    On Error Resume Next
    s = CStr(doc.Variables(STATE_VAR_NAME).Value)
    On Error GoTo 0
    ReadStateVar = UCase$(s)
End Function

' =====================================================================
' Helper: write the state variable to a document
' =====================================================================
Private Sub WriteStateVar(ByVal doc As Document, ByVal val As String)
    On Error Resume Next
    doc.Variables(STATE_VAR_NAME).Value = val
    If Err.Number <> 0 Then
        Err.Clear
        doc.Variables.Add Name:=STATE_VAR_NAME, Value:=val
    End If
    On Error GoTo 0
End Sub

' =====================================================================
' SUITE A  StripTrailingCR  (local reimplementation)
' =====================================================================
Private Sub Suite_StripTrailingCR()
    Debug.Print ""
    Debug.Print "--- Suite_StripTrailingCR ---"

    ' A1: empty string is unchanged
    AssertEqual LocalStripCR(""), "", _
        "A1 empty string returns empty string"

    ' A2: string without any CR is unchanged
    AssertEqual LocalStripCR("hello"), "hello", _
        "A2 string without CR unchanged"

    ' A3: single trailing CR is stripped
    AssertEqual LocalStripCR("hello" & Chr(13)), "hello", _
        "A3 single trailing CR stripped"

    ' A4: string that IS only Chr(13) becomes empty
    AssertEqual LocalStripCR(Chr(13)), "", _
        "A4 lone CR becomes empty string"

    ' A5: double trailing CR — only the LAST one is stripped
    AssertEqual LocalStripCR("hello" & Chr(13) & Chr(13)), _
                "hello" & Chr(13), _
        "A5 double trailing CR: only one stripped (last)"

    ' A6: LF (Chr 10) is NOT treated as a CR
    AssertEqual LocalStripCR("hello" & Chr(10)), "hello" & Chr(10), _
        "A6 trailing LF (Chr 10) is not stripped"

    ' A7: mid-string CR is not affected
    AssertEqual LocalStripCR("a" & Chr(13) & "b"), "a" & Chr(13) & "b", _
        "A7 mid-string CR is not stripped"

    ' A8: Arabic text with trailing CR
    AssertEqual LocalStripCR(Chr(1589) & Chr(1583) & Chr(1585) & Chr(13)), _
                Chr(1589) & Chr(1583) & Chr(1585), _
        "A8 Arabic text with trailing CR stripped correctly"

    ' A9: trailing CRLF — only the CR at the end matters; the sequence
    '     ends with Chr(10), so nothing is stripped
    AssertEqual LocalStripCR("text" & Chr(13) & Chr(10)), _
                "text" & Chr(13) & Chr(10), _
        "A9 trailing CRLF: LF is last char so nothing stripped"
End Sub

' =====================================================================
' SUITE B  Separator-parsing (sadr / ajuz split)
' =====================================================================
Private Sub Suite_SeparatorParsing()
    Debug.Print ""
    Debug.Print "--- Suite_SeparatorParsing ---"

    Dim sadr As String, ajuz As String
    Dim sepPos As Long

    ' B1: typical line with text on both sides
    sepPos = InStr(1, "sadr ** ajuz", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "sadr ** ajuz", sepPos, sadr, ajuz
    AssertEqual sadr, "sadr", "B1 left side parsed correctly"
    AssertEqual ajuz, "ajuz", "B1 right side parsed correctly"

    ' B2: ** at very start (empty sadr after Trim)
    sepPos = InStr(1, "** ajuz", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "** ajuz", sepPos, sadr, ajuz
    AssertEqual sadr, "", "B2 separator at start: sadr is empty"
    AssertEqual ajuz, "ajuz", "B2 separator at start: ajuz correct"

    ' B3: ** at very end (empty ajuz after Trim)
    sepPos = InStr(1, "sadr **", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "sadr **", sepPos, sadr, ajuz
    AssertEqual sadr, "sadr", "B3 separator at end: sadr correct"
    AssertEqual ajuz, "", "B3 separator at end: ajuz is empty"

    ' B4: leading/trailing spaces around ** are trimmed
    sepPos = InStr(1, "  sadr  **  ajuz  ", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "  sadr  **  ajuz  ", sepPos, sadr, ajuz
    AssertEqual sadr, "sadr", "B4 sadr trimmed"
    AssertEqual ajuz, "ajuz", "B4 ajuz trimmed"

    ' B5: Arabic text split
    Dim arabic As String
    arabic = Chr(1589) & Chr(1583) & Chr(1585) & " ** " & Chr(1593) & Chr(1580) & Chr(1586)
    sepPos = InStr(1, arabic, POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit arabic, sepPos, sadr, ajuz
    AssertEqual sadr, Chr(1589) & Chr(1583) & Chr(1585), "B5 Arabic sadr (صدر)"
    AssertEqual ajuz, Chr(1593) & Chr(1580) & Chr(1586), "B5 Arabic ajuz (عجز)"

    ' B6: multiple ** in line — InStr returns first occurrence
    sepPos = InStr(1, "a ** b ** c", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "a ** b ** c", sepPos, sadr, ajuz
    AssertEqual sadr, "a", "B6 multiple ** - sadr uses first occurrence"
    AssertEqual ajuz, "b ** c", "B6 multiple ** - ajuz contains remainder"

    ' B7: exactly "**" with no spaces (both sides empty after Trim)
    sepPos = InStr(1, "**", POETRY_SEPARATOR, vbBinaryCompare)
    LocalSplit "**", sepPos, sadr, ajuz
    AssertEqual sadr, "", "B7 bare ** - sadr empty"
    AssertEqual ajuz, "", "B7 bare ** - ajuz empty"

    ' B8: separator not present (sepPos = 0) — guard condition check
    sepPos = InStr(1, "no separator here", POETRY_SEPARATOR, vbBinaryCompare)
    Assert (sepPos = 0), "B8 no separator: InStr returns 0 (guard fires)"
End Sub

' =====================================================================
' SUITE C  State management  (GetMode / SetMode via document variable)
' =====================================================================
Private Sub Suite_StateManagement()
    Debug.Print ""
    Debug.Print "--- Suite_StateManagement ---"

    Dim doc As Document
    Set doc = NewTempDoc()

    ' C1: document variable absent → state reads as empty (not "ON")
    AssertEqual ReadStateVar(doc), "", _
        "C1 no variable set: state is empty string"
    Assert (ReadStateVar(doc) <> "ON"), _
        "C1 no variable set: state is not ON"

    ' C2: write MODE_ON ("ON") → reads back as "ON"
    WriteStateVar doc, "ON"
    AssertEqual ReadStateVar(doc), "ON", _
        "C2 after writing ON: reads back ON"

    ' C3: write MODE_OFF ("") → reads back as ""
    WriteStateVar doc, ""
    AssertEqual ReadStateVar(doc), "", _
        "C3 after writing empty string: reads back empty"

    ' C4: overwrite existing variable (no duplicate-add error)
    WriteStateVar doc, "ON"
    WriteStateVar doc, "ON"  ' second write to same variable
    AssertEqual ReadStateVar(doc), "ON", _
        "C4 double write ON: no error, reads ON"

    ' C5: state value is case-normalised to upper-case on read
    '     (UCase$ is applied in GetMode; simulate with lower-case write)
    WriteStateVar doc, "on"
    AssertEqual UCase$(ReadStateVar(doc)), "ON", _
        "C5 lower-case 'on' in variable: UCase normalises to ON"

    CloseTempDoc doc
End Sub

' =====================================================================
' SUITE D  FormatArabicPoetryOnEnter — no separator present
' =====================================================================
Private Sub Suite_FormatOnEnter_NoSeparator()
    Debug.Print ""
    Debug.Print "--- Suite_FormatOnEnter_NoSeparator ---"

    Dim doc As Document
    Set doc = NewTempDoc()
    doc.Activate

    ' Position cursor at end of document
    Selection.EndKey Unit:=wdStory

    ' Type a line that does NOT contain **
    Selection.TypeText Text:="plain text without separator"
    FormatArabicPoetryOnEnter

    ' D1: no table should have been created
    AssertLong doc.Tables.Count, 0, _
        "D1 no separator: no table created"

    ' D2: a new paragraph was inserted (cursor moved to next line)
    Assert (Selection.Start > 1), _
        "D2 no separator: Enter still creates new paragraph"

    CloseTempDoc doc
End Sub

' =====================================================================
' SUITE E  FormatArabicPoetryOnEnter — separator present
' =====================================================================
Private Sub Suite_FormatOnEnter_WithSeparator()
    Debug.Print ""
    Debug.Print "--- Suite_FormatOnEnter_WithSeparator ---"

    Dim doc As Document
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory

    Selection.TypeText Text:="sadr ** ajuz"
    FormatArabicPoetryOnEnter

    ' E1: exactly one table was created
    AssertLong doc.Tables.Count, 1, _
        "E1 separator present: exactly one table created"

    If doc.Tables.Count = 1 Then
        Dim tbl As Table
        Set tbl = doc.Tables(1)

        ' E2: table dimensions
        AssertLong tbl.Rows.Count, 1, _
            "E2 table has 1 row"
        AssertLong tbl.Columns.Count, 2, _
            "E2 table has 2 columns"

        ' E3: cell content — sadr in cell(1,1), ajuz in cell(1,2)
        Dim cellText1 As String, cellText2 As String
        cellText1 = LocalStripCR(tbl.Cell(1, 1).Range.Text)
        cellText2 = LocalStripCR(tbl.Cell(1, 2).Range.Text)

        AssertEqual cellText1, "sadr", _
            "E3 cell(1,1) contains sadr (text before **)"
        AssertEqual cellText2, "ajuz", _
            "E3 cell(1,2) contains ajuz (text after **)"
    End If

    CloseTempDoc doc
End Sub

' =====================================================================
' SUITE F  FormatArabicPoetryOnEnter — cursor already inside a table
' =====================================================================
Private Sub Suite_FormatOnEnter_InTable()
    Debug.Print ""
    Debug.Print "--- Suite_FormatOnEnter_InTable ---"

    Dim doc As Document
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory

    ' Insert an ordinary table and move into it
    Dim existingTbl As Table
    Set existingTbl = doc.Tables.Add(Selection.Range, 1, 2)
    existingTbl.Cell(1, 1).Range.Select

    ' Type a line with ** inside the table cell
    Selection.TypeText Text:="a ** b"
    FormatArabicPoetryOnEnter

    ' F1: the handler must NOT create an extra table when already in a table
    '     (the guard  If wasInTable Then Exit Sub  fires)
    AssertLong doc.Tables.Count, 1, _
        "F1 cursor in table: no new table created (guard fires)"

    CloseTempDoc doc
End Sub

' =====================================================================
' SUITE G  Table structure checks
' =====================================================================
Private Sub Suite_TableStructure()
    Debug.Print ""
    Debug.Print "--- Suite_TableStructure ---"

    Dim doc As Document
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory

    Selection.TypeText Text:="right ** left"
    FormatArabicPoetryOnEnter

    If doc.Tables.Count <> 1 Then
        mFailed = mFailed + 1
        Debug.Print "  [FAIL]  G-SETUP: table not created; skipping structure tests"
        CloseTempDoc doc
        Exit Sub
    End If

    Dim tbl As Table
    Set tbl = doc.Tables(1)

    ' G1: all six border line styles are wdLineStyleNone
    Assert (tbl.Borders(wdBorderTop).LineStyle = wdLineStyleNone), _
        "G1 border top is none"
    Assert (tbl.Borders(wdBorderBottom).LineStyle = wdLineStyleNone), _
        "G1 border bottom is none"
    Assert (tbl.Borders(wdBorderLeft).LineStyle = wdLineStyleNone), _
        "G1 border left is none"
    Assert (tbl.Borders(wdBorderRight).LineStyle = wdLineStyleNone), _
        "G1 border right is none"
    Assert (tbl.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone), _
        "G1 border horizontal is none"
    Assert (tbl.Borders(wdBorderVertical).LineStyle = wdLineStyleNone), _
        "G1 border vertical is none"

    ' G2: table preferred width is 100 %
    Assert (tbl.PreferredWidthType = wdPreferredWidthPercent), _
        "G2 table preferred-width type is percent"
    AssertLong CLng(tbl.PreferredWidth), 100, _
        "G2 table preferred width is 100%"

    ' G3: each column is 50 %
    Assert (tbl.Columns(1).PreferredWidthType = wdPreferredWidthPercent), _
        "G3 column 1 preferred-width type is percent"
    AssertLong CLng(tbl.Columns(1).PreferredWidth), 50, _
        "G3 column 1 preferred width is 50%"
    Assert (tbl.Columns(2).PreferredWidthType = wdPreferredWidthPercent), _
        "G3 column 2 preferred-width type is percent"
    AssertLong CLng(tbl.Columns(2).PreferredWidth), 50, _
        "G3 column 2 preferred width is 50%"

    ' G4: reading order is RTL
    Assert (tbl.Rows(1).Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl), _
        "G4 row reading order is RTL"
    Assert (tbl.Cell(1, 1).Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl), _
        "G4 cell(1,1) reading order is RTL"
    Assert (tbl.Cell(1, 2).Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl), _
        "G4 cell(1,2) reading order is RTL"

    CloseTempDoc doc
End Sub

' =====================================================================
' SUITE H  Edge cases
' =====================================================================
Private Sub Suite_EdgeCases()
    Debug.Print ""
    Debug.Print "--- Suite_EdgeCases ---"

    Dim doc As Document

    ' H1: ** at start of line — sadr is empty, ajuz has content
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="** ajuz only"
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H1 separator at start: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "", _
            "H1 cell(1,1) is empty (no sadr)"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "ajuz only", _
            "H1 cell(1,2) contains ajuz text"
    End If
    CloseTempDoc doc

    ' H2: ** at end of line — ajuz is empty, sadr has content
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="sadr only **"
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H2 separator at end: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "sadr only", _
            "H2 cell(1,1) contains sadr text"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "", _
            "H2 cell(1,2) is empty (no ajuz)"
    End If
    CloseTempDoc doc

    ' H3: bare ** with no surrounding text — both cells empty
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="**"
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H3 bare **: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "", _
            "H3 cell(1,1) empty"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "", _
            "H3 cell(1,2) empty"
    End If
    CloseTempDoc doc

    ' H4: multiple ** — first occurrence is the split point
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="a ** b ** c"
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H4 multiple **: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "a", _
            "H4 cell(1,1) = text before first **"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "b ** c", _
            "H4 cell(1,2) = remainder after first **"
    End If
    CloseTempDoc doc

    ' H5: spaces around ** are trimmed
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="  left  **  right  "
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H5 spaces around **: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "left", _
            "H5 cell(1,1) sadr trimmed"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "right", _
            "H5 cell(1,2) ajuz trimmed"
    End If
    CloseTempDoc doc

    ' H6: Arabic text round-trip
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Dim arabicLine As String
    arabicLine = Chr(1589) & Chr(1583) & Chr(1585) & _
                 " ** " & _
                 Chr(1593) & Chr(1580) & Chr(1586)  ' "صدر ** عجز"
    Selection.TypeText Text:=arabicLine
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H6 Arabic text: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), _
                    Chr(1589) & Chr(1583) & Chr(1585), _
            "H6 cell(1,1) = Arabic sadr"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), _
                    Chr(1593) & Chr(1580) & Chr(1586), _
            "H6 cell(1,2) = Arabic ajuz"
    End If
    CloseTempDoc doc

    ' H7: paragraph with only spaces (no **) — no table
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="   "
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 0, "H7 spaces-only paragraph: no table"
    CloseTempDoc doc

    ' H8: empty paragraph (no text at all) — no table
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    FormatArabicPoetryOnEnter  ' Enter on a blank paragraph
    AssertLong doc.Tables.Count, 0, "H8 empty paragraph: no table"
    CloseTempDoc doc

    ' H9 (regression): ** contained inside a longer word should still trigger
    '    e.g. "foo**bar" — the separator need not have surrounding spaces
    Set doc = NewTempDoc()
    doc.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="foo**bar"
    FormatArabicPoetryOnEnter
    AssertLong doc.Tables.Count, 1, "H9 embedded **: table created"
    If doc.Tables.Count = 1 Then
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 1).Range.Text), "foo", _
            "H9 cell(1,1) = text before embedded **"
        AssertEqual LocalStripCR(doc.Tables(1).Cell(1, 2).Range.Text), "bar", _
            "H9 cell(1,2) = text after embedded **"
    End If
    CloseTempDoc doc
End Sub
