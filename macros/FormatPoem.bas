Attribute VB_Name = "FormatPoem"
' =====================================================================
' FormatPoem - Arabic Poetry Auto-Formatter (first-step / minimal)
' ---------------------------------------------------------------------
' On Enter, if the just-finished paragraph contains "**", split it
' into a borderless 2-column RTL table:
'   right cell  = text BEFORE  "**"   (الصدر)
'   left cell   = text AFTER   "**"   (العجز)
' The reading order is RTL, so "right" / "left" here refer to the
' visual position of the cell on screen / page.
'
' Toggle the feature with  ToggleArabicPoetryTable . The toggle stores
' its state in a document variable so AutoOpen can re-bind Enter when
' Word restarts.
'
' Install:
'   1. Open the VBA editor (Alt+F11).
'   2. File > Import File... and pick this .bas file.
'   3. Save Normal.dotm (so the binding persists).
'   4. Run  ToggleArabicPoetryTable  once to turn the feature ON.
'
' Macros only run in macro-enabled files (.docm / .doc / .dotm /
' Normal.dotm itself).
' =====================================================================

Option Explicit

Private Const POETRY_SEPARATOR As String = "**"
Private Const STATE_VAR_NAME   As String = "ArabicPoetryTableMode"
Private Const MODE_ON          As String = "ON"
Private Const MODE_OFF         As String = ""

Private mBound As Boolean

' =====================================================================
' Public toggle
' =====================================================================
Public Sub ToggleArabicPoetryTable()
    If GetMode() = MODE_ON Then
        SetMode MODE_OFF
        UnbindEnter
        MsgBox "Arabic Poetry (Table) is OFF." & vbCrLf & _
               "Enter behaves normally.", _
               vbInformation, "Arabic Poetry Formatter"
    Else
        SetMode MODE_ON
        BindEnter "FormatArabicPoetryOnEnter"
        MsgBox "Arabic Poetry (Table) is ON." & vbCrLf & _
               "Type a line with ** between the two hemistichs and " & _
               "press Enter.", _
               vbInformation, "Arabic Poetry Formatter"
    End If
End Sub

Public Sub AutoExec()
    If GetMode() = MODE_ON Then BindEnter "FormatArabicPoetryOnEnter"
End Sub

Public Sub AutoOpen()
    AutoExec
End Sub

' =====================================================================
' Enter handler
' =====================================================================
Public Sub FormatArabicPoetryOnEnter()
    Dim wasInTable As Boolean
    Dim previousPara As Paragraph
    Dim prevText As String
    Dim sepPos As Long
    Dim adjacentTbl As Table

    wasInTable = Selection.Information(wdWithInTable)
    Selection.TypeParagraph
    If wasInTable Then Exit Sub

    On Error Resume Next
    Set previousPara = Selection.Range.Paragraphs(1).Previous
    On Error GoTo 0
    If previousPara Is Nothing Then Exit Sub

    prevText = StripTrailingCR(previousPara.Range.Text)
    sepPos = InStr(1, prevText, POETRY_SEPARATOR, vbBinaryCompare)
    If sepPos <= 0 Then Exit Sub

    Application.ScreenUpdating = False
    Set adjacentTbl = AdjacentPoetryTable(previousPara)
    If adjacentTbl Is Nothing Then
        ConvertLineToPoetryTable previousPara.Range, prevText, sepPos
    Else
        AppendLineToPoetryTable adjacentTbl, previousPara, prevText, sepPos
    End If
    Application.ScreenUpdating = True
End Sub

' =====================================================================
' Core: convert one paragraph into a 2-column borderless RTL table
' =====================================================================
Private Sub ConvertLineToPoetryTable(ByVal rngLine As Range, _
                                     ByVal lineText As String, _
                                     ByVal sepPos As Long)
    Dim sadr As String, ajuz As String
    Dim sepLen As Long
    Dim insertRange As Range
    Dim tbl As Table

    sepLen = Len(POETRY_SEPARATOR)
    sadr = Trim$(Left$(lineText, sepPos - 1))
    ajuz = Trim$(Mid$(lineText, sepPos + sepLen))

    ' Wipe the paragraph's text but keep its trailing CR, then insert
    ' the table at the resulting empty position.
    Set insertRange = rngLine.Duplicate
    If insertRange.End > insertRange.Start Then _
        insertRange.End = insertRange.End - 1
    If insertRange.End > insertRange.Start Then _
        insertRange.Text = ""
    insertRange.Collapse Direction:=wdCollapseStart

    Set tbl = rngLine.Document.Tables.Add( _
        Range:=insertRange, _
        NumRows:=1, _
        NumColumns:=2, _
        DefaultTableBehavior:=wdWord9TableBehavior, _
        AutoFitBehavior:=wdAutoFitWindow)

    With tbl
        .Borders.Enable = False
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone

        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Columns(1).PreferredWidthType = wdPreferredWidthPercent
        .Columns(1).PreferredWidth = 50
        .Columns(2).PreferredWidthType = wdPreferredWidthPercent
        .Columns(2).PreferredWidth = 50

        ' Cell margins: bottom = 0.2 cm (others left at Word defaults)
        .BottomPadding = Application.CentimetersToPoints(0.2)

        .Rows(1).Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    End With

    ' RTL: cell(1,1) is visually on the right (sadr), cell(1,2) on the left (ajuz).
    FillCell tbl.Cell(1, 1), sadr
    FillCell tbl.Cell(1, 2), ajuz
End Sub

Private Sub FillCell(ByVal c As Cell, ByVal text As String)
    Dim r As Range
    Set r = c.Range
    r.Collapse Direction:=wdCollapseStart
    r.Text = text
    c.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    c.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    c.VerticalAlignment = wdCellAlignVerticalCenter
End Sub

' =====================================================================
' Append-vs-create dispatch helpers
' ---------------------------------------------------------------------
' If a 2-column poetry table sits above the just-finished `**`
' paragraph -- with at most MAX_GAP empty paragraphs between them --
' the new hemistichs become an extra row in that table instead of a
' separate one, and the in-between blank paragraphs are eaten.
' =====================================================================
Private Function AdjacentPoetryTable(ByVal para As Paragraph) As Table
    Const MAX_GAP As Long = 4
    Dim cur As Paragraph
    Dim gap As Long
    Dim t As Table

    On Error Resume Next
    Set cur = para.Previous
    On Error GoTo 0

    gap = 0
    Do While Not cur Is Nothing
        If cur.Range.Information(wdWithInTable) Then
            On Error Resume Next
            Set t = cur.Range.Tables(1)
            On Error GoTo 0
            If t Is Nothing Then Exit Function
            If t.Columns.Count <> 2 Then Exit Function
            Set AdjacentPoetryTable = t
            Exit Function
        End If

        If Len(StripTrailingCR(cur.Range.Text)) > 0 Then Exit Function

        gap = gap + 1
        If gap > MAX_GAP Then Exit Function

        On Error Resume Next
        Set cur = cur.Previous
        On Error GoTo 0
    Loop
End Function

Private Sub AppendLineToPoetryTable(ByVal tbl As Table, _
                                    ByVal para As Paragraph, _
                                    ByVal lineText As String, _
                                    ByVal sepPos As Long)
    Dim sadr As String, ajuz As String
    Dim sepLen As Long
    Dim existingAlign As Long
    Dim newRow As Row
    Dim slab As Range

    sepLen = Len(POETRY_SEPARATOR)
    sadr = Trim$(Left$(lineText, sepPos - 1))
    ajuz = Trim$(Mid$(lineText, sepPos + sepLen))

    ' Capture the existing top-row alignment so a previously snugged
    ' table (wdAlignParagraphDistribute) keeps it on the new row.
    existingAlign = tbl.Rows(1).Range.ParagraphFormat.Alignment

    Set newRow = tbl.Rows.Add

    FillCell newRow.Cells(1), sadr
    FillCell newRow.Cells(2), ajuz

    newRow.Range.ParagraphFormat.Alignment = existingAlign

    ' Eat any 1-4 blank paragraphs between the (now-bigger) table and
    ' the user's `**` line, plus the line itself. Building the slab
    ' from the LIVE tbl.Range.End / para.Range.End AFTER Rows.Add
    ' avoids the auto-expanding-Range bug where a Range captured
    ' before Rows.Add silently grew over the new row.
    Set slab = para.Range.Document.Range( _
        Start:=tbl.Range.End, _
        End:=para.Range.End)
    slab.Delete
End Sub

' =====================================================================
' Manual: snug the table's left/right cell margins to the widest text
' ---------------------------------------------------------------------
' For the table at the cursor, find the largest left/right cell padding
' (Word's "default cell margins") that doesn't force any cell's text
' onto a new line, then back off by 0.1 cm per side so the widest text
' sits in its cell with a ~0.2 cm horizontal halo. The same padding
' applies to every cell, so narrower-text cells gain visibly more
' breathing room.
'
' If there isn't room for the 0.2 cm halo, the halo is dropped and the
' margins go right up to the wrap boundary. If a cell already wraps at
' zero padding, its natural vertical text span at zero padding is taken
' as that cell's baseline so the pre-existing wrap doesn't keep the
' bisection pinned at zero.
'
' Implementation: bisect the padding value, and after each step ask
' Word's own layout engine whether any cell's text now spans more
' vertical distance than it did at zero padding. The signal is the
' vertical gap between the cursor at the cell text's logical start
' and its logical end: ~0 for one display line, grows by ~one line
' height per wrap. This avoids fragile horizontal pixel-width
' measurements that are easy to get wrong for RTL / centered Arabic
' text, and avoids ComputeStatistics(wdStatisticLines), which on a
' per-cell range counts paragraph-delimited lines (always 1 here)
' rather than display lines.
' =====================================================================
Public Sub AdjustPoetryTableMargins()
    Const HALO_PER_SIDE_CM As Single = 0.1
    Const TOL_CM           As Single = 0.01
    Const MAX_ITER         As Long = 30
    Const WRAP_EPS_PTS     As Single = 1!

    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Place the cursor inside a poetry table first.", _
               vbExclamation, "Arabic Poetry Formatter"
        Exit Sub
    End If

    Dim tbl As Table
    Set tbl = Selection.Tables(1)

    Dim origLeft As Single, origRight As Single
    Dim origAlign As Long
    origLeft = tbl.LeftPadding
    origRight = tbl.RightPadding
    origAlign = tbl.Range.ParagraphFormat.Alignment

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ' Upper bound: half the narrowest cell. Padding above this would
    ' give that cell a non-positive content area.
    Dim cel As Cell
    Dim hiPts As Single
    hiPts = 1E+30
    For Each cel In tbl.Range.Cells
        If cel.Width / 2 < hiPts Then hiPts = cel.Width / 2
    Next cel
    If hiPts <= 0 Then GoTo Cleanup

    ' Capture each cell's natural text vertical span at zero padding.
    tbl.LeftPadding = 0
    tbl.RightPadding = 0

    Dim cellCount As Long
    cellCount = tbl.Range.Cells.Count
    If cellCount = 0 Then GoTo Cleanup

    Dim baseline() As Single
    ReDim baseline(1 To cellCount)

    Dim hasAnyText As Boolean
    Dim idx As Long
    idx = 0
    For Each cel In tbl.Range.Cells
        idx = idx + 1
        baseline(idx) = CellTextSpan(cel)
        If CellHasText(cel) Then hasAnyText = True
    Next cel

    If Not hasAnyText Then
        ' No text to size around. Restore and exit silently.
        tbl.LeftPadding = origLeft
        tbl.RightPadding = origRight
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Bisect for the largest padding where no cell exceeds its baseline.
    Dim lo As Single, hi As Single, mid As Single
    Dim tolPts As Single
    Dim iter As Long
    lo = 0
    hi = hiPts
    tolPts = Application.CentimetersToPoints(TOL_CM)

    For iter = 1 To MAX_ITER
        If (hi - lo) <= tolPts Then Exit For
        mid = (lo + hi) / 2
        tbl.LeftPadding = mid
        tbl.RightPadding = mid
        If TableExceedsBaseline(tbl, baseline, WRAP_EPS_PTS) Then
            hi = mid
        Else
            lo = mid
        End If
    Next iter

    ' lo = largest padding that didn't introduce new wrapping.
    ' Back off by 0.1 cm per side to leave a 0.2 cm total halo,
    ' unless there isn't room for it.
    Dim haloPts As Single
    haloPts = Application.CentimetersToPoints(HALO_PER_SIDE_CM)

    Dim finalPad As Single
    If lo >= haloPts Then
        finalPad = lo - haloPts
    Else
        finalPad = lo
    End If
    If finalPad < 0 Then finalPad = 0

    tbl.LeftPadding = finalPad
    tbl.RightPadding = finalPad

    ' Distribute-justify every cell (Ctrl+Shift+J) so each hemistich
    ' spreads evenly across the now-snug content area.
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphDistribute

    Application.ScreenUpdating = True
    Exit Sub

Cleanup:
    On Error Resume Next
    tbl.LeftPadding = origLeft
    tbl.RightPadding = origRight
    tbl.Range.ParagraphFormat.Alignment = origAlign
    Application.ScreenUpdating = True
End Sub

Private Function TableExceedsBaseline(ByVal tbl As Table, _
                                      ByRef baseline() As Single, _
                                      ByVal eps As Single) As Boolean
    Dim cel As Cell
    Dim idx As Long
    Dim span As Single
    idx = 0
    For Each cel In tbl.Range.Cells
        idx = idx + 1
        span = CellTextSpan(cel)
        If span > baseline(idx) + eps Then
            TableExceedsBaseline = True
            Exit Function
        End If
    Next cel
    TableExceedsBaseline = False
End Function

Private Function CellTextSpan(ByVal c As Cell) As Single
    ' Vertical gap (points) between the cursor at the cell text's
    ' logical start and at its logical end. ~0 when the text fits on
    ' one display line; grows by ~one line height for each wrap.
    Dim r As Range
    Dim ch As String
    Dim rs As Range, re As Range
    Dim y1 As Single, y2 As Single

    Set r = c.Range.Duplicate
    Do While r.End > r.Start
        ch = Right$(r.Text, 1)
        If ch = vbCr Or ch = Chr$(7) Then
            r.End = r.End - 1
        Else
            Exit Do
        End If
    Loop

    If r.End <= r.Start Then
        CellTextSpan = 0
        Exit Function
    End If

    Set rs = r.Duplicate
    rs.Collapse Direction:=wdCollapseStart
    Set re = r.Duplicate
    re.Collapse Direction:=wdCollapseEnd

    y1 = rs.Information(wdVerticalPositionRelativeToPage)
    y2 = re.Information(wdVerticalPositionRelativeToPage)

    ' Information returns -1 when the position can't be determined.
    If y1 < 0 Or y2 < 0 Then
        Err.Raise vbObjectError + 513, "CellTextSpan", _
                  "Unable to resolve cell layout position."
    End If

    CellTextSpan = Abs(y2 - y1)
End Function

Private Function CellHasText(ByVal c As Cell) As Boolean
    Dim s As String
    Dim ch As String
    s = c.Range.Text
    Do While Len(s) > 0
        ch = Right$(s, 1)
        If ch = vbCr Or ch = Chr$(7) Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    CellHasText = (Len(Trim$(s)) > 0)
End Function

' =====================================================================
' Enter binding (Normal.dotm)
' =====================================================================
Private Sub BindEnter(ByVal macroName As String)
    If mBound Then Exit Sub

    Dim kc As Long
    Dim kb As KeyBinding
    kc = Application.BuildKeyCode(wdKeyReturn)

    On Error Resume Next
    CustomizationContext = NormalTemplate
    Set kb = KeyBindings.Key(kc)
    If Not kb Is Nothing Then
        If kb.Command <> macroName Then kb.Clear
    End If
    KeyBindings.Add KeyCode:=kc, _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:=macroName
    On Error GoTo 0

    mBound = True
End Sub

Private Sub UnbindEnter()
    Dim kc As Long
    Dim kb As KeyBinding
    kc = Application.BuildKeyCode(wdKeyReturn)

    On Error Resume Next
    CustomizationContext = NormalTemplate
    Set kb = KeyBindings.Key(kc)
    If Not kb Is Nothing Then
        If InStr(kb.Command, "FormatArabicPoetryOnEnter") > 0 Then kb.Clear
    End If
    On Error GoTo 0

    mBound = False
End Sub

' =====================================================================
' State + helpers
' =====================================================================
Private Function GetMode() As String
    Dim s As String
    s = ""
    On Error Resume Next
    s = CStr(ActiveDocument.Variables(STATE_VAR_NAME).Value)
    Err.Clear
    On Error GoTo 0
    GetMode = UCase$(s)
End Function

Private Sub SetMode(ByVal newMode As String)
    On Error Resume Next
    Err.Clear
    ActiveDocument.Variables(STATE_VAR_NAME).Value = newMode
    If Err.Number <> 0 Then
        Err.Clear
        ActiveDocument.Variables.Add Name:=STATE_VAR_NAME, Value:=newMode
    End If
    On Error GoTo 0
End Sub

Private Function StripTrailingCR(ByVal s As String) As String
    If Len(s) > 0 Then
        If Right$(s, 1) = Chr(13) Then
            StripTrailingCR = Left$(s, Len(s) - 1)
            Exit Function
        End If
    End If
    StripTrailingCR = s
End Function
