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
    ConvertLineToPoetryTable previousPara.Range, prevText, sepPos
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
