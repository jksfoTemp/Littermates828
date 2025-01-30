Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

' Remember to save (version) often (source control is another
' conversation completely) incrementing the # when you do and
' annotating the notes in the second sheet
'
' Tip - try to keep the characters to fewer than 80 per line,
' the ' _' characters can be used to extend a line
'
' Tip - Make sure the 'Insert' key is not active
'
' Tip - Keep the numeric keypad on
'
' Tip - One can usually put a carriage return in the text of a
' cell but hitting [Alt] and the enter key
'
' Tip - In the VBA editor on the lower left of the active code
' window there (should be) are two buttons that look like one
' paragraph and then two paragraphs. Keep the second one selected
' or the code that is not in you current block (Sub or Function)
' will be hidden. An absolutely stupid 'feature' for an IDE
' (integrated development environment - tool that one uses to
' write code, often has features like syntax highlighting, error
' checking, formatting, commenting and other)
'
' Tip - Because this is (really, for instance, no line numbers) such a shitty editor I often
' mark where I was like 'zzzzzzz
'
' Tip - if starting in a new(ly installed) editor be sure to turn
' off syntax checking or it will drive you fucking nuts. Syntax will be
' checked when you run vs. when you type. ([Tools] [Options]
' [Auto Syntax Check])
'
' Tip - Do what I say not what I do
'
' Tip - to see the properties of an object, like a button, go to the
' "Developer" menu and select "Design Mode", remember to turn it off
' when you want them to actually work
'
' Tip - don't leave it open on tab 2 or 3 or the contents might be wiped out
'
' Tip - F9 for a breakpoint, F8 to step through the code
'
' Tip - Prolly too many calls to save routine but whateva
'
' Tip - Protect the sheets ([Review] [Protect Sheet]) for "Revisions"
' and "Data" to avoid messing with them accidentally. Annoying but
' worth it
'
'
'
'
'
'
'  TODO - See sheet 2
'  1.) Reset button postions
'  2.) Naming issue, existing card, SaveAs error
'  2.5) Save at end and reset focus
'  9.) Scroll to card two for drawing and then for words, reset focus after
'  3.) Check 'New Card' vs. 'Stub for Open'?
'  18.) Negative distance on button spread
'  17.) Place timer such that buttons paint in a more responsive manner
'  8.) Casino effect on/off
'  7.) Custom form for entry
'  5.) New categories with checked
'  6.) Already selected, watch out for recursion
'  4.) Revision log
'  9.) Game selection
'30.) Clean, explain sheet two
'31.) Note about developer ribbon change
'32.) Note about xlsm and change to .txt for emailing


Global gParentWindowFileName As String
Global gParentWindowSheetName As String
Global gChildWindowFileName As String
Global gChildWindowSheetName As String
Global gTrace As Boolean
Global gLastRowInDomain As Integer
Global gDelay As Integer
Global gCaller As String
Global gGreeting As String
Global gGreetingEnd As String
Global gSecondGameCardName
' This must have been some interrupted thought, I'm sure it seemed
' like a good idea at the time. Ha.
Global gSleep As Integer
Global gMaxLoopValueForWritingGrid As Integer ' 26
Global gLazyLoopSkipTop As Integer
Global gLazyLoopSkipBottom As Integer
Global gInput As String
Global gMaxRecursion As Integer

#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'#If VBA7 Then
'    ' Alias "ShellExecuteA"
'    Declare PtrSafe Function ShellExecute Lib "shell32.dll" () _
'        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#Else
'    ' Alias "ShellExecuteA"
'    Declare Function ShellExecute Lib "shell32.dll" _
'        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#End If

Sub mySetGlobals()
  Dim myLocalChildRand
  Dim lSecondGameCardName

  gCaller = "mySetGlobals()"
  Call mySlimeFile(gCaller)

  On Error GoTo Ack

    gTrace = True
    gDelay = 2 '  0 ' 20
    If (gGreeting = "") Then
      gGreeting = "Therapy"
      gGreetingEnd = " Bingo!"
      gGreeting = gGreeting & gGreetingEnd
    End If
    gParentWindowFileName = ActiveWorkbook.Name
    gParentWindowSheetName = ActiveSheet.Name

    'Init the generator
    Randomize
    myLocalChildRand = rnd(1)
    gChildWindowSheetName = "Card_" & myLocalChildRand
    gChildWindowFileName = gChildWindowSheetName & ".xlsx"

    lSecondGameCardName = myLocalChildRand + 0.2
    'lSecondGameCardName = "Card_" + lSecondGameCardName
    gSecondGameCardName = "Card_" + Str(lSecondGameCardName)

    gMaxLoopValueForWritingGrid = 26
    gLazyLoopSkipTop = 2
    gLazyLoopSkipBottom = 25
    gSleep = 100 ' 250 ' 500
    gCaller = ""
    gMaxRecursion = 26

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myWelcome0()
  gCaller = "myWelcome0()"
  Call mySlimeFile(gCaller)
  Call mySlimeTrail(gCaller)
  'gGreeting = "Therapy Bingo!"
  On Error GoTo Ack

    MsgBox ("Welcome to " & gGreeting)
    'MsgBox ("Welcome to Testing!")
    Call myBegin
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

'#If VBA7 Then
'    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
'        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#Else
'    ''Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
'        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#End If


Sub myBegin()
  gCaller = "myBegin()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack

    'TODO - BREAK OUT SET GLOBALS TO ONE STEP.
    Call mySetGlobals

    Call myStubForOpen
    Call myClearSheetValues
    '' Call mySetup ' Replaced
    Call mySetWindowAndSheetDimensions
    'DoEvents

    'Data manipulation
    Call mySort
    Call myFindLastCellWithValue
    Call myClearFormats
    Call myColorListOfWords
    Call myTrim
    Call myRemoveDupes

    'Eh, don't really care
    'Call myCheckSpelling

    'Get words and check for dupes
    Call mySelectWords
    DoEvents
    Call myDrawNewCard1
    DoEvents
    'Sleep (gSleep)
    'ShellExecuteA
    'ShellExecuteA.Sleep (gSleep)
    Sleep (gSleep)
    Call myDrawNewCard2
    DoEvents

    'Sleep (gSleep) ' in ms

    'Assign card values - clean extra calls
    Call myCopySourceValuesToDestination1
    Sleep (gSleep)
    Call myCopySourceValuesToDestination2


    ' Not Needed - Call Save
    'Call myPrint
    'Call myHideButtons(0)
    'Call myStubForOpen

    Call myTidyUp

  Exit Sub
Ack:
    Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myTidyUp()
  gCaller = "myTidyUp()"
  Call mySlimeTrail(gCaller)
  'Dim lEnablePrint As Boolean
  Dim lSPrinter As String

  On Error GoTo Ack
    ActiveSheet.cmdOpen.Enabled = True
    ActiveSheet.cmdNew.Enabled = True
    ActiveSheet.cmdColor.Enabled = True

    Range("A1").Select
    ActiveWorkbook.Save

    If Application.ActivePrinter <> "" Then
      ActiveSheet.cmdPrint.Enabled = True
    Else
      ActiveSheet.cmdPrint.Enabled = False
    End If

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)

End Sub

'Function myIsPrinter() As Boolean
'  gCaller = "myIsPrinter()"
'  Call mySlimeTrail(gCaller)
'  Dim p As Object
'
'  On Error GoTo Ack
'    p = GetObject("winmgmts:\\.\root\CIMV2")
'
'    myIsPrinter() = False
'  Exit Function
'Ack:
'  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
'End Function

''Used?
'Sub mySetup()
'  gCaller = "mySetup() - Replaced."
'  Call mySlimeTrail(gCaller)
'
'  On Error GoTo Ack
''    Randomize
''    gParentWindowSheetName = ActiveSheet.name
''    gParentWindowFileName = ActiveWorkbook.name
'  Exit Sub
'Ack:
'    Debug.Print ("Error - " & gCaller & " : " & Err.Description)
'End Sub

''kkkkkkk

Sub myCopySourceValuesToDestination1()
  gCaller = "myCopySourceValuesToDestination1()"
  Call mySlimeTrail(gCaller)

  Dim mySourceWorkbook 'As Excel.Workbook
  Dim mySourceWorksheet 'As Excel.Worksheet
  Dim myChildWorkbook 'As Excel.Workbook
  Dim myChildWorksheet 'As Excel.Worksheet
  Dim sRow As Integer, sColumn As Integer, dRow As Integer, dColumn As Integer
  Dim myVal As String

  On Error GoTo Ack
    ' Set objects
    Set mySourceWorkbook = Workbooks(gParentWindowFileName)
    mySourceWorkbook.Activate
    Set mySourceWorksheet = Excel.Worksheets(gParentWindowSheetName)

    Set myChildWorkbook = Workbooks(gChildWindowFileName)
    myChildWorkbook.Activate

    'Set myChildWorksheet = Worksheets(gChildWindowSheetName)
    Set myChildWorksheet = Worksheets("Sheet1")

    sRow = 1
    sColumn = 3
    dRow = 3
    dColumn = 1

    mySourceWorkbook.Activate

    While sRow < gMaxLoopValueForWritingGrid 'This should mpt be hard coded
      'ick infinite loop
      If (sRow < 2) Or (sRow > (gMaxLoopValueForWritingGrid - 1)) Then
        ' Only log this if it is the beginning or the end
        Debug.Print ("myCopySourceValuesToDestination1() iteration: While sRow = " & sRow)
      End If

      'myVal = ActiveSheet.Cells(sRow, sColumn)
      myVal = ActiveSheet.Cells(sRow, sColumn + 1)
      myChildWorksheet.Cells.WrapText = True
      myChildWorksheet.Cells(dRow, dColumn) = myVal

      If sRow Mod 5 = 0 Then
        dColumn = dColumn + 1
        dRow = 2
      End If
        ' Play / casino
        ' This value should not be hard coded
      Call myButtonColors(sRow, gDelay)
      sRow = sRow + 1
      dRow = dRow + 1
      'Call myButtonColors
    Wend

    myChildWorkbook.Save
    mySourceWorkbook.Activate
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

' Note that this is a cheat from v1, code should not be copy/pasted
' rather it should be reused via parameterization
Sub myCopySourceValuesToDestination2()
  gCaller = "myCopySourceValuesToDestination2()"
  Call mySlimeTrail(gCaller)

  Dim mySourceWorkbook 'As Excel.Workbook
  Dim mySourceWorksheet 'As Excel.Worksheet
  Dim myChildWorkbook 'As Excel.Workbook
  Dim myChildWorksheet 'As Excel.Worksheet
  Dim sRow As Integer, sColumn As Integer, dRow As Integer, dColumn As Integer
  Dim myVal As String

  On Error GoTo Ack

    ' Set objects
    Set mySourceWorkbook = Workbooks(gParentWindowFileName)
    mySourceWorkbook.Activate
    Set mySourceWorksheet = Excel.Worksheets(gParentWindowSheetName)

    Set myChildWorkbook = Workbooks(gChildWindowFileName)
    myChildWorkbook.Activate
    ActiveWindow.ScrollRow = 15
    ' JK Fix this new
    'Set myChildWorksheet = Worksheets(gChildWindowSheetName)
    Dim ick As String
    ' what is /where set
    ick = gChildWindowSheetName
    'Debug.Print ("Errrrrrrrrrrrrrrrrrrrror: " & gChildWindowSheetName)
    Set myChildWorksheet = Worksheets("Sheet1")

    sRow = 1
    sColumn = 5
    dRow = 17
    dColumn = 1

    mySourceWorkbook.Activate

    While sRow < gMaxLoopValueForWritingGrid
      If (sRow < 2) Or (sRow > (gMaxLoopValueForWritingGrid - 1)) Then
        ' Only log this if it is the beginning or the end
        Debug.Print ("myCopySourceValuesToDestination2() iteration: While sRow = " & sRow)
      End If
      ' myVal = ActiveSheet.Cells(sRow, sColumn)
      myVal = ActiveSheet.Cells(sRow, sColumn + 1)
      myChildWorksheet.Cells(dRow, dColumn) = myVal

      If sRow Mod 5 = 0 Then
        dColumn = dColumn + 1
        dRow = 16
      End If
      ' Play / casino
      ' This value should not be hard coded
      Call myButtonColors(sRow, gDelay)
      sRow = sRow + 1
      dRow = dRow + 1
      'Call myButtonColors
    Wend

    myChildWorkbook.Activate
    ActiveWindow.ScrollRow = 1
    'Might be unnecessary ...
    ActiveSheet.Cells(1, 1).Select

    myChildWorkbook.Save
    mySourceWorkbook.Activate
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myPrint()
  gCaller = "myPrint()"
  Call mySlimeTrail(gCaller)

  Dim myChildWorkbook 'As Excel.Workbook
  Dim myChildWorksheet 'As Excel.Worksheet

  On Error GoTo Ack
    Call myHideButtons(1)

    Set myChildWorkbook = Workbooks(gChildWindowFileName)
    myChildWorkbook.Activate
    myChildWorkbook.Save
    Set myChildWorksheet = Worksheets(gChildWindowSheetName)
    myChildWorksheet.Activate
    ActiveSheet.PageSetup.CenterHorizontally = True
    ActiveSheet.PrintOut Copies:=1

    Call myHideButtons(1)

    'Optional
    ActiveSheet.cmdPrint.Enabled = False
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myDrawNewCard1()
  gCaller = "myDrawNewCard1()"
  Call mySlimeTrail(gCaller)

  Dim newCard As Workbook
  Dim lTLCorner, lBRCorner
  Dim lFoo As String

  On Error GoTo Ack
    'Set newCard = Workbooks(gChildWindowSheetName)
    Set newCard = Workbooks(gChildWindowFileName)

    newCard.Activate
    ActiveWindow.DisplayGridlines = False
    newCard.Save
    'myChildWorkbook.Save
    'newCard.SaveAs (gChildWindowFileName)
    Range("A1").Activate

    'Get rid of other sheets
    Call myFixSheets(newCard)

    'Row 1
    lTLCorner = "A1"
    lBRCorner = "E1"
    Call mySetMergedContent(gChildWindowSheetName, lTLCorner, lBRCorner, gGreeting, 14, 1, 1)
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 0)
    'newCard.Save

    lFoo = gChildWindowSheetName + " - A whole day of fun! "

    'Row 2
    lTLCorner = "A2"
    lBRCorner = "E2"
    Call mySetMergedContent(gChildWindowSheetName, lTLCorner, lBRCorner, lFoo, 10, 0, 0)
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 0)
    'newCard.Save

    'Playing grid
    lTLCorner = "A3"
    lBRCorner = "E7"
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 2)
    Call myICheatAndSetTheDamnGrid(gChildWindowSheetName, lTLCorner, lBRCorner)

    'Dark game board border
    lTLCorner = "A3"
    lBRCorner = "E7"
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    'newCard.Save

    ' Cheat - first row taller
    ActiveSheet.Rows(1).RowHeight = 40

    ' Cheat - columns A-E width = 16, height = 24
    ActiveSheet.Range("A3:F7").ColumnWidth = 16
    ActiveSheet.Range("A3:F7").RowHeight = 40
    ActiveSheet.Range("A3:F7").HorizontalAlignment = xlCenter
    ActiveSheet.Range("A3:F7").VerticalAlignment = xlCenter
    'ActiveSheet.range("A3:F7").WordWrap = True

    'Test - refocus
    'Call myButtonColors
    ActiveSheet.Cells(1, 1).Select
    newCard.Save

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myDrawNewCard2()
  gCaller = "myDrawNewCard2()"
  Call mySlimeTrail(gCaller)

  Dim newCard As Workbook
  Dim lTLCorner, lBRCorner
  Dim lSecondGameCardName As String, bar

  ' Yes, a lazy cheat, copy / paste of #1

  On Error GoTo Ack
    lSecondGameCardName = gChildWindowSheetName
    lSecondGameCardName = Replace(lSecondGameCardName, "Card_", "")
    'Should not be needed but just in case ...
    lSecondGameCardName = Replace(lSecondGameCardName, ".xlsx", "")
    lSecondGameCardName = Replace(lSecondGameCardName, ".xlsm", "")
    lSecondGameCardName = lSecondGameCardName + 0.2
    gSecondGameCardName = "Card_" + lSecondGameCardName + " - A whole day of fun! "

    Set newCard = Workbooks(gChildWindowFileName)
    newCard.Activate
    Range("A15").Select
    ActiveWindow.ScrollRow = 15
    'Range("A15").ScrollRow

    'Row 1
    lTLCorner = "A15"
    lBRCorner = "E15"
    Call mySetMergedContent(gChildWindowSheetName, lTLCorner, lBRCorner, gGreeting, 14, 1, 1)
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 0)

    'Row 2
    lTLCorner = "A16"
    lBRCorner = "E16"
    Call mySetMergedContent(gChildWindowSheetName, lTLCorner, lBRCorner, gSecondGameCardName, 10, 0, 0)
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 0)

    'Playing grid
    lTLCorner = "A17"
    lBRCorner = "E21"
    Call mySetIntBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 2)
    Call myICheatAndSetTheDamnGrid(gChildWindowSheetName, lTLCorner, lBRCorner)

    'Dark game board border
    lTLCorner = "A17"
    lBRCorner = "E21"
    Call mySetExtBorders(gChildWindowSheetName, lTLCorner, lBRCorner, 3)
    'newCard.Save

    ' Cheat - first row taller
    ActiveSheet.Rows(15).RowHeight = 40

    ' Cheat - columns A-E width = 16, heigh = 24
    ActiveSheet.Range("A17:F21").ColumnWidth = 16
    ActiveSheet.Range("A17:F21").RowHeight = 40
    ActiveSheet.Range("A17:F21").HorizontalAlignment = xlCenter
    ActiveSheet.Range("A17:F21").VerticalAlignment = xlCenter
    'ActiveSheet.range("A17:F21").WordWrap = True

    'Test - refocus
    'Call myButtonColors
    ActiveSheet.Cells(1, 1).Select
    newCard.Save

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySetExtBorders(lTgt, lStart, lEnd, lWgt)
  gCaller = "mySetExtBorders(" & lTgt & ", " & lStart & ", " & lEnd & ", " & lWgt & ")"
  Call mySlimeTrail(gCaller)

  Dim lWeight As Variant

  On Error GoTo Ack
    'Weights:
    '   0 - none
    '   1 - thin
    '   2 - normal ' Default
    '   3 - bold
    Select Case lWgt
        Case 0
            lWeight = xlNone
        Case 1
            lWeight = xlThin
        Case 3
            lWeight = xlThick
        Case Else
            lWeight = xlMedium
    End Select

    Windows(gChildWindowFileName).Activate
    Range(lStart, lEnd).Select

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = lWeight
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = lWeight
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = lWeight
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = lWeight
    End With
    'Call myButtonColors
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySetIntBorders(lTgt, lStart, lEnd, lWgt)
  gCaller = "mySetIntBorders(" & lTgt & ", " & lStart & ", " & lEnd & ", " & lWgt & ")"
  Call mySlimeTrail(gCaller)

  Dim lWeight As Variant 'As String

  On Error GoTo Ack
    'Weights:
    '   0 - none
    '   1 - thin
    '   2 - normal ' Default
    '   3 - bold
    Select Case lWgt
        Case 0
            lWeight = xlNone
        Case 1
            lWeight = xlThin
        Case 3
            lWeight = xlThick
        Case Else
            lWeight = 3 'xlMedium
    End Select

    Windows(gChildWindowFileName).Activate
    Range(lStart, lEnd).Select

    Selection.Borders(xlInsideVertical).LineStyle = lWeight
    Selection.Borders(xlInsideHorizontal).LineStyle = lWeight
    Selection.Borders(xlDiagonalDown).LineStyle = lWeight
    Selection.Borders(xlDiagonalUp).LineStyle = lWeight
    'Call myButtonColors
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySetMergedContent(lTgt, lStart, lEnd, lText, lSize, lBold, lItalic)
  gCaller = "mySetMergedContent(" & lTgt & ", " & lStart & ", " & lEnd & ", " & lText & ", " & lSize & ", " & lBold & ", " & lItalic & ")"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    Windows(gChildWindowFileName).Activate
    Range(lStart, lEnd).Select

    With Selection
      ' Most assumed, can parameterize
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
    End With

    With Selection.Font
      ' Most assumed, can parameterize
      .Size = lSize
      .Italic = lItalic
      .Bold = lBold
      .Name = "Calibri"
      .Strikethrough = False
      .Superscript = False
      .Subscript = False
      .OutlineFont = False
      .Shadow = False
      .Underline = xlUnderlineStyleNone
      .ThemeColor = xlThemeColorLight1
      .TintAndShade = 0
      .ThemeFont = xlThemeFontMinor
    End With

    ActiveCell.FormulaR1C1 = lText ' gGreeting
    'Call myButtonColors
  Exit Sub
Ack:
  Debug.Print (gCaller & Err.Description)
End Sub

Sub myFixSheets(tgt As Workbook)
  gCaller = "myFixSheets(" & tgt.Name & " As Workbook)"
  Call mySlimeTrail(gCaller)

  Dim mySheetname As String
  'Dim wb As Workbook
  Dim ws As Worksheet

  On Error GoTo Ack
    mySheetname = Replace(tgt.Name, ".xlsx", "")
    mySheetname = "Sheet1"
    'Figure out a way to remove the  other sheets
    '   WHY THIS FUCKING DANCE????????????????
    Application.DisplayAlerts = False

    'Check sheet existence
'    If tgt.Sheets(3).name = "Sheet3" Then
'        tgt.Sheets(3).Delete
'    End If
'    If tgt.Sheets(2).name = "Sheet2" Then
'        tgt.Sheets(2).Delete
'    End If
'
    'Better, should be function, lazy
    'If wb Is Nothing Then
    'Set wb = ActiveWorkbook '.Worksheets 'ThisWorkbook
    'foreach ws in wb.Sheets
     For Each ws In ActiveWorkbook.Worksheets
   '   Set sht = wb.Sheets(shtName)
      If ws.Name <> "Sheet1" Then
        ws.Delete
      End If
      Next

    Application.DisplayAlerts = True

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myColorListOfWords()
  gCaller = "myColorListOfWords()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    Range("A1:A" & gLastRowInDomain).Select
    Selection.Font.Color = RGB(20, 55, 222)
    ActiveWorkbook.Save
    Range("A1").Select
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySort()
  gCaller = "mySort()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    'Set Parent window as active
    Workbooks(gParentWindowFileName).Activate
    'ActiveWorkbook.Worksheets(1).Columns(1).Select
    'Workbooks(1).Activate
    'gParentWindowSheetName
    ActiveWorkbook.Worksheets(1).Cells(1, 1).Select

    With ActiveWorkbook.Worksheets(1).Sort
    'With ActiveWorkbook.Worksheets(gParentWindowSheetName).Sort
      .SetRange Columns(1)
      .Header = xlNo
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
    End With
    ActiveWorkbook.Worksheets(1).Cells(1, 1).Select
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myFindLastCellWithValue()
  gCaller = "myFindLastCellWithValue()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    gLastRowInDomain = Cells(Rows.Count, 1).End(xlUp).row
    Debug.Print ("There are " & gLastRowInDomain & " rows in the domain (gLastRowInDomain)")
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myTrim()
  gCaller = "myTrim()"
  Call mySlimeTrail(gCaller)

  Dim r

  On Error GoTo Ack
    r = 1
    For r = 1 To gLastRowInDomain
        Cells(r, 1) = Trim(Cells(r, 1))
    Next r
  Exit Sub
Ack:
  Debug.Print (gCaller & Err.Description)
End Sub

Sub myClearFormats()
  gCaller = "myClearFormats()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    Range("A1", "A" & gLastRowInDomain).ClearFormats
  Exit Sub
Ack:
  Debug.Print (gCaller & Err.Description)
End Sub

Sub myRemoveDupes()
  gCaller = "myRemoveDupes()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    Range("A1", "A" & gLastRowInDomain).RemoveDuplicates Columns:=1, Header:=xlNo
    Columns("A:A").Sort key1:=Range("A1"), order1:=xlAscending, Header:=xlNo
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myCheckSpelling()
  gCaller = "myCheckSpelling()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
  ' zzzzzzzzzzzzzzzzz activate/switch to first workbook
  'ActiveWorkbook.
  'gParentWindowSheetName
  'Set mySourceWorkbook = Workbooks(gParentWindowFileName)
    Workbooks(gParentWindowFileName).Activate
    ActiveWorkbook.Worksheets(gParentWindowSheetName).Activate

    Application.DisplayAlerts = False

    ActiveSheet.Range("A1:A" & gLastRowInDomain).Select
    Selection.CheckSpelling
    ActiveWorkbook.Worksheets(gParentWindowSheetName).Range("A1:A1").Select
    Application.DisplayAlerts = True

    Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myICheatAndSetTheDamnGrid(lTgt, lStart, lEnd)
  gCaller = "myICheatAndSetTheDamnGrid(" & lTgt & ", " & lStart & ", " & lEnd & ")"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack
    Windows(gChildWindowFileName).Activate
    Range(lStart, lEnd).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    With Selection.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
    End With
  Exit Sub
Ack:
    Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySetWindowAndSheetDimensions()
  gCaller = "mySetWindowAndSheetDimensions()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack

    'Application
    ActiveWindow.WindowState = xlNormal
    Application.Width = 480
    Application.Height = 440
    Application.Top = 0
    Application.Left = 0

    'Control window 'Redundant
    'ActiveWindow.Width = 640
    'ActiveWindow.Height = 540
    'ActiveWindow.Top = 10
    'ActiveWindow.Left = 10

    'Set column widths, just not working right, cheat
    'Columns("A:A").Select
    'Columns("A:A").EntireColumn.AutoFit
    'Columns("A:A").Width = Columns("A:A").Width + 80
    'Columns("B:B").Select
    'Columns("B:B").EntireColumn.AutoFit
    'Columns("B:B").Width = Columns("B:B").Width + 80
    ' ? Which is the active file book?
    ' Be declarative and are the G vars set yet?
    Columns("A:B").AutoFit

    Workbooks.Add
    'ActiveWorkbook.SaveAs gChildWindowSheetName
    'zzzzzzzzzz Correct
    'ActiveWorkbook.SaveAs gChildWindowFileName
    ActiveWorkbook.SaveAs gChildWindowFileName

    ActiveWindow.Width = Application.Width + 60 ' 500
    'ActiveWindow.Height = 300 ' 1 Game
    ActiveWindow.Height = 520 ' 2 Games
    ActiveWindow.Top = 20
    ActiveWindow.Left = 500

    Range("A1:A1").Select
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myFindAndRecordNewWord(lIteration, _
  targetColumnIndex, targetColumnWord)
  gCaller = "myFindNewWord()"
  Call mySlimeTrail(gCaller)
  On Error GoTo Ack

'TODO: Check if any of these are used
  Dim iter As Integer, lRandomIndex As Integer
  Dim lTempVal, lTempValTest As String, lMarkMe As String
  Dim lRnd

  iter = 1
  lMarkMe = "Selected"

  While (iter < gMaxRecursion) ' 26
    lRnd = rnd(1)
    lRandomIndex = Int(gLastRowInDomain * lRnd) + 1
    lTempVal = Cells(lRandomIndex, 1).Value
    lTempValTest = Cells(lRandomIndex, 2).Value

    If (lTempValTest <> "") Then
      ' Already selected
      lRandomIndex = lRandomIndex + 1
      iter = iter - 1
    ElseIf (Cells(lRandomIndex, 1).Value = "") Then
      ' If we got here then we ran out of values, start at the
      ' top of the list. Potential for a loop if the column is blank but …
      lRandomIndex = 1
      iter = iter - 1
    Else
      ' Really handling the display should be separate
      ' (Separation Of Concerns) Mark it as selected
      Cells(lRandomIndex, 2).Value = lMarkMe
      ' Record it
      Cells(iter, targetColumnIndex).Value = lRandomIndex
      Cells(iter, targetColumnWord).Value = lTempVal
      Cells(lRandomIndex, 1).Font.Bold = True
    End If

    ' Play time
    Call myButtonColors(iter, gDelay)

    iter = iter + 1 ' Avoid endless loop
  Wend

 'Just not working well, cheat
  'Columns(targetColumnIndex).Select
  'Columns(targetColumnIndex).EntireColumn.AutoFit
  'Columns(targetColumnIndex).ColumnWidth = 50 'Columns(targetColumnIndex).Width + 80
  'Columns(targetColumnWord).Select
  'Columns(targetColumnWord).EntireColumn.AutoFit
  'Columns(targetColumnWord).ColumnWidth = 100 ' Columns(targetColumnWord).Width + 80
  Columns("A:F").AutoFit

    '    Else ' We have exceed 25 attempts looking for an unreserved value
    '      lErr = "Error - recursion exceeded. " & _
    '        "gCaller: " & gCaller & _
    '        "returnMyWord: " & returnMyWord & _
    '        "lRandomIndex: " & lRandomIndex & _
    '        "lIteration: " & lIteration & _
    '        "iter25: " & iter25 & _
    '        "iter: " & iter & _
    '        "lTempVal: " & lTempVal & _
    '        "lTempValTest: " & lTempValTest
    '      MsgBox (lErr)
    '      Debug.Print (lErr)
    '    End If
Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub mySelectWords()
  gCaller = "mySelectWords()"
  Call mySlimeTrail(gCaller)

 Dim iter As Integer

  On Error GoTo Ack
    ' First iteration
    Call myFindAndRecordNewWord(iter, 3, 4)

    ' Second iteration
    Call myFindAndRecordNewWord(iter, 5, 6)

    Columns("B:F").EntireColumn.AutoFit
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myStubForOpen()
  gCaller = "myStubForOpen()"
  Call mySlimeTrail(gCaller)
  Dim myH As Integer, myW As Integer, myL As Integer, myT As Integer

  On Error GoTo Ack
    myH = 25
    myW = 100
    myL = 20
    myT = 20

    If gParentWindowFileName <> "" Then
      Windows(gParentWindowFileName).Activate
      ActiveWindow.DisplayGridlines = False
    End If

    ActiveSheet.cmdOpen.Visible = True
    ActiveSheet.cmdNew.Visible = True
    ActiveSheet.cmdPrint.Visible = True
    ActiveSheet.cmdColor.Visible = True

    ' rrrrrrrrrrrrrrrrrrrrrrrrreset
    ActiveSheet.cmdOpen.Enabled = True
    ActiveSheet.cmdNew.Enabled = False 'True
    ActiveSheet.cmdPrint.Enabled = False 'False
    ActiveSheet.cmdColor.Enabled = False 'True

    ActiveSheet.cmdOpen.BackColor = RGB(248, 252, 155)
    ActiveSheet.cmdNew.BackColor = RGB(248, 252, 155)
    ActiveSheet.cmdPrint.BackColor = RGB(248, 252, 155)
    ActiveSheet.cmdColor.BackColor = RGB(248, 252, 155)

    ActiveSheet.cmdOpen.Top = myT * 1
    ActiveSheet.cmdNew.Top = myT * 3
    ActiveSheet.cmdPrint.Top = myT * 5
    ActiveSheet.cmdColor.Top = myT * 7

    ActiveSheet.cmdOpen.Height = myH
    ActiveSheet.cmdNew.Height = myH
    ActiveSheet.cmdPrint.Height = myH
    ActiveSheet.cmdColor.Height = myH

    ActiveSheet.cmdOpen.Width = myW
    ActiveSheet.cmdNew.Width = myW
    ActiveSheet.cmdPrint.Width = myW
    ActiveSheet.cmdColor.Width = myW

    ActiveSheet.cmdOpen.Left = myL
    ActiveSheet.cmdNew.Left = myL
    ActiveSheet.cmdPrint.Left = myL
    ActiveSheet.cmdColor.Left = myL

    DoEvents

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myClearSheetValues()
  gCaller = "myClearSheetValues()"
  Call mySlimeTrail(gCaller)

  'gParentWindowSheetName = ActiveSheet.name
  Dim lCheckActiveSheet As String

  lCheckActiveSheet = ActiveSheet.Name

  ' Activesheet.  . name

  On Error GoTo Ack
    ' put in check here that the first sheet is selected
    If lCheckActiveSheet = gParentWindowSheetName Then
      ' Stooped but for some reason this has fucked me more than once
      If lCheckActiveSheet = "Hello!" Then
        Columns("B:F").Value = ""
      End If
    End If
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

'''''zzzzzzzzzzzzz

Sub myHideButtons(onOff)
  gCaller = "myHideButtons(" & onOff & ")"
  Call mySlimeTrail(gCaller)
  Dim mySourceWorkbook

  On Error GoTo Ack
    Set mySourceWorkbook = Workbooks(gParentWindowFileName)
    mySourceWorkbook.Activate

    If (onOff = 1) Then
        'ActiveSheet.cmdOpen.Visible = False
        'ActiveSheet.cmdNew.Visible = False
        'ActiveSheet.cmdPrint.Visible = False
        'ActiveSheet.cmdColor.Visible = False
        ActiveSheet.cmdOpen.Enabled = False
        ActiveSheet.cmdNew.Enabled = True
        ActiveSheet.cmdPrint.Enabled = False
        ActiveSheet.cmdColor.Enabled = False
    Else
        ActiveSheet.cmdOpen.Visible = True
        ActiveSheet.cmdNew.Visible = True
        ActiveSheet.cmdPrint.Visible = True
        ActiveSheet.cmdColor.Visible = True
        ActiveSheet.cmdOpen.Enabled = True
        ActiveSheet.cmdNew.Enabled = True
        ActiveSheet.cmdPrint.Enabled = True
        ActiveSheet.cmdColor.Enabled = True
    End If
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

Sub myColorAndReset()
  gCaller = "myColorAndReset()"
  Call mySlimeTrail(gCaller)

  On Error GoTo Ack

    Call myStubForOpen

    ' zzzzzzzzz needed?
    'ActiveSheet.cmdColor.Left = 20
    ActiveSheet.Cells(1, 1).Select

  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

' play casino
Sub myButtonColors(iter, delay)
  gCaller = "myButtonColors(" & iter & ", " & delay & ")"

  ' iter test here
  If ((iter < 2) Or (iter > (gMaxLoopValueForWritingGrid - 1))) Then
    Call mySlimeTrail(gCaller)
  End If

  Dim r0 As Integer, r1 As Integer, r2 As Integer, myWhichOne As Integer, _
    myWhichWay As Integer, rTop As Integer, rLeft As Integer

  On Error GoTo Ack
    r0 = Int((255) * rnd + 1)
    r1 = Int((255) * rnd + 1)
    r2 = Int((255) * rnd + 1)
    rTop = Int((10) * rnd + 1)
    rLeft = Int((10) * rnd + 1)
    myWhichOne = r0 Mod 4
    myWhichWay = r0 Mod 2
    ' zzzzzzzzzzzzzzzzz  myButtonColors(1, 20) : Object doesn't support this property or method

    ' todo - did I want to add in a negative variance here? can't remember ...

    'Toopid
    'Windows(gChildWindowFileName).Activate
    Windows(gParentWindowFileName).Activate

    'Turned of button movement
    Select Case myWhichOne
      Case 0
        ActiveSheet.cmdOpen.BackColor = RGB(r0, r1, r2)
        'ActiveSheet.cmdOpen.Top = ActiveSheet.cmdOpen.Top + (rTop * myWhichWay)
        'ActiveSheet.cmdOpen.Left = ActiveSheet.cmdOpen.Left + (rLeft * myWhichWay)
      Case 1
        ActiveSheet.cmdNew.BackColor = RGB(r0, r1, r2)
        'ActiveSheet.cmdNew.Top = ActiveSheet.cmdNew.Top + (rTop * myWhichWay)
        'ActiveSheet.cmdNew.Left = ActiveSheet.cmdNew.Left + (rLeft * myWhichWay)
      Case 2
        ActiveSheet.cmdPrint.BackColor = RGB(r0, r1, r2)
        'ActiveSheet.cmdPrint.Top = ActiveSheet.cmdPrint.Top + (rTop * myWhichWay)
        'ActiveSheet.cmdPrint.Left = ActiveSheet.cmdPrint.Left + (rLeft * myWhichWay)
      Case Else
        ActiveSheet.cmdColor.BackColor = RGB(r0, r1, r2)
        'ActiveSheet.cmdColor.Top = ActiveSheet.cmdColor.Top + (rTop * myWhichWay)
        'ActiveSheet.cmdColor.Left = ActiveSheet.cmdColor.Left + (rLeft * myWhichWay)
    End Select

    ' why hardcoded? turns out it was an old hack
    ' ActiveSheet.cmdOpen.Left = 40
    Sleep (delay)
    DoEvents
  Exit Sub
Ack:
  Debug.Print ("Error - " & gCaller & " : " & Err.Description)
End Sub

'
Sub mySlimeFile(arg As String)
  On Error GoTo Ack:
    If gTrace = True Then
      ' JK: 10.22.1630
      '  Conacatenating all this together just leads to syntax hell, gotta love BarfBA
      Debug.Print _
        "gParentWindowFileName: '" & gParentWindowFileName & "' | " & "gParentWindowSheetName: '" & gParentWindowSheetName & "' | " & vbNewLine & Spc(4) & "gChildWindowFileName: '" & gChildWindowFileName & "' | " & "gChildWindowSheetName: '" & gChildWindowSheetName & "' | " & vbNewLine & Spc(4) & "gTrace: '" & gTrace & "' | " & "gLastRowInDomain: '" & gLastRowInDomain & "' | " & "gDelay: '" & gDelay & "' | "

'      Debug.Print _
 '       "gParentWindowFileName: '" & gParentWindowFileName & "' | " _
  '    & "gParentWindowSheetName: '" & gParentWindowSheetName & "' | " _
   '   & vbNewLine & Spc(4) & _
    '      "gChildWindowFileName: '" & gChildWindowFileName & "' | " _
     '   & "gChildWindowSheetName: '" & gChildWindowSheetName & "' | " _
      '  & vbNewLine; Spc(4); _
       '   "gTrace: '" & gTrace & "' | " _
        '& "gLastRowInDomain: '" & gLastRowInDomain & "' | " _
        '& "gDelay: '" & gDelay & "' | ";


  End If
Exit Sub
Ack:
  Debug.Print ("Ha! Error in the error routine: " & Err.Description)
End Sub

Sub mySlimeTrail(arg As String)
  On Error GoTo Ack:

    Debug.Print (gCaller & " | " & "arg: /" & arg & "/ | ")
    'Debug.Print "    " & Err.Description; '& vbNewLine
    Debug.Print "    " & Err.Description '& vbNewLine

Exit Sub
Ack:
  ' Debug.Print ("Ha! Error in the error routine: " & Err.Description)
  Debug.Print ("Ha! Error in the error routine: " & arg)

End Sub


