Sub MegAlyse()
' Paul Beverley - Version 20.01.22
' Launches a selected series of analysis macros

' Works with: AccentAlyse, CapitAlyse, CenturyAlyse, DocAlyse,
' FullNameAlyse, HyphenAlyse, ListAlyse, ProperNounAlyse,
' SpecialSortsLister, SpellAlyse, SpellingErrorListerBilingual,
' WordPairAlyse

myAlyses = "ListAlyse, DocAlyse"
myAlyses = "SpellAlyse, WordPairAlyse"
myAlyses = "DocAlyse, HyphenAlyse, ProperNounAlyse, SpellAlyse, WordPairAlyse"

saveResultsFiles = False

' On a Mac, you will need something like:
myFolder = "/Users/Paul/My Documents/Macro_stuff/"

' On Windows, you will need something like:
myFolder = "C:\Documents and Settings\Paul\Macro_stuff\"

myResponse = MsgBox("MegAlyse" & vbCr & vbCr & _
     "Run:      " & myAlyses & "?", vbQuestion _
     + vbYesNoCancel, "MegAlyse")
If myResponse <> vbYes Then Exit Sub

' Don't change this filename
tempFile = myFolder & "zzTestFile"
stTime = Time
thisLanguage = Selection.LanguageID

Set rng = ActiveDocument.Content
Documents.Add
Set testFile = ActiveDocument

Selection.FormattedText = rng.FormattedText
Selection.EndKey Unit:=wdStory
If ActiveDocument.Endnotes.count > 0 Then
  Set thisDocRange = testFile.Content
  thisDocRange.Collapse wdCollapseEnd
  thisDocRange.FormattedText = _
       testFile.StoryRanges(wdEndnotesStory).FormattedText
End If
If ActiveDocument.Footnotes.count > 0 Then
  Set thisDocRange = testFile.Content
  thisDocRange.Collapse wdCollapseEnd
  thisDocRange.FormattedText = _
       testFile.StoryRanges(wdFootnotesStory).FormattedText
End If

' copy all the textboxes to the end of the text
shCount = testFile.Shapes.count
If shCount > 0 Then
  Selection.EndKey Unit:=wdStory
  Selection.TypeText Text:=vbCr & vbCr
  For j = 1 To shCount
    Set shp = ActiveDocument.Shapes(j)
    If shp.Type <> 24 And shp.Type <> 3 Then
      If shp.TextFrame.HasText Then
        Set rng = shp.TextFrame.TextRange
        Selection.FormattedText = rng.FormattedText
        Selection.EndKey Unit:=wdStory
      End If
    End If
  Next
  For j = shCount To 1 Step -1
    ActiveDocument.Shapes(j).Delete
  Next
End If
Set rng = ActiveDocument.Content
rng.Fields.Unlink
rng.Revisions.AcceptAll
For Each myPic In ActiveDocument.InlineShapes
  myPic.Delete
Next myPic
ActiveDocument.Content.LanguageID = thisLanguage
ActiveDocument.SaveAs FileName:=tempFile

myAlyses = Replace("," & myAlyses & ",", ",,", ",")
myAlyses = Replace(myAlyses, " ", "")
thisArray = Split(myAlyses, ",")
For i = 1 To UBound(thisArray) - 1
  rprt = thisArray(i) & " started?????: " & Left(Time, 5) & vbCr
  Debug.Print rprt
  Application.Run MacroName:=thisArray(i)
  DoEvents
Next i
rprt = vbCr & "Finished at: " & Left(Time, 5)
Debug.Print rprt

If saveResultsFiles Then
  For Each myDoc In Documents
    myName = myDoc.Name
    If Left(myName, 8) = "Document" Then
      Set rng = myDoc.Content
      newName = Left(rng.Text, 40)
      crPos = InStr(newName, vbCr)
      If crPos > 3 Then
        newName = Left(newName, crPos - 1)
        myDoc.Activate
        myFullFilename = myFolder & newName
        ActiveDocument.SaveAs FileName:=myFullFilename
      End If
    End If
  Next myDoc
End If

testFile.Activate
ActiveDocument.Close SaveChanges:=False
Beep
myTime = Timer
Do
Loop Until Timer > myTime + 0.2
Beep
End Sub

Sub AccentAlyse()
' Paul Beverley - Version 06.02.20
' Analyses all the words that contain an accent

' These are the accents to watch out for
allAccents = "áÁàÀâÂäÄÃãÅåçÇéÉèÈêÊëËíÍìÌîÎñÑóÓòÒôÔöÖõÕøØßúÚùÙûÛüÜýÝÿŸ"

' For including 'Central European' characters
addExtraCharacters = True
fromUnicode = 256
toUnicode = 382

' Minimum word length
minLength = 3

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)
If doingSeveralMacros = False Then
  myResponse = MsgBox("    AccentAlyse" & vbCr & vbCr & _
       "Analyse this document?", vbQuestion _
       + vbYesNoCancel, "AccentAlyse")
  If myResponse <> vbYes Then Exit Sub
End If
If addExtraCharacters = True Then
  For i = fromUnicode To toUnicode
    allAccents = allAccents & ChrW(i)
  Next i
End If

myLead = " . . . "
Set rng = ActiveDocument.Content
Documents.Add
Set resultDoc = ActiveDocument
Set res = ActiveDocument.Content
Documents.Add
Set testDoc = ActiveDocument
Selection.Text = rng.Text
Selection.HomeKey Unit:=wdStory
Selection.TypeText "Finding accented words..." & vbCr _
     & vbCr & vbCr
Selection.Start = 0
Selection.range.Style = ActiveDocument.Styles(wdStyleHeading1)
Selection.Font.Underline = True

Set rng = testDoc.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[" & allAccents & "]"
  .Font.Underline = False
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute
End With

Do While rng.Find.Found = True
  DoEvents
  rng.Expand wdWord
  myEnd = rng.End
  If InStr(" " & ChrW(8217), Right(rng.Text, 1)) _
       > 0 Then rng.End = rng.End - 1
  If InStr(" " & ChrW(8217), Right(rng.Text, 1)) _
       > 0 Then rng.End = rng.End - 1
  myAccWord = rng.Text
  If Len(myAccWord) >= minLength Then
    Set rng2 = testDoc.Content
    myTot = rng2.End
    DoEvents
    With rng2.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = myAccWord
      .Wrap = wdFindContinue
      .MatchWholeWord = True
      .Replacement.Text = "^&!"
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
    myCount = testDoc.Content.End - myTot
    WordBasic.EditUndo
    DoEvents
    Selection.End = 0
    With rng2.Find
      .Replacement.Text = "^&"
      .Replacement.Font.Underline = True
      .MatchWholeWord = True
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
    DoEvents
    
    ' Find an accent-different word
    findWord = ""
    myAltWord = ""
    For i = 1 To Len(myAccWord)
      myChar = Mid(myAccWord, i, 1)
      If InStr(allAccents, myChar) > 0 Then
        findWord = findWord & "^$"
      Else
        findWord = findWord & myChar
      End If
    Next i
    testDoc.Content.Font.StrikeThrough = False
    With rng2.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = findWord
      .Font.Underline = False
      .Font.StrikeThrough = False
      .Wrap = wdFindStop
      .Replacement.Text = ""
      .Forward = True
      .MatchWildcards = False
      .Execute
    End With
    Do While rng2.Find.Found = True
      DoEvents
      myAltWord = rng2.Text
      Set rng3 = testDoc.Content
      myTot = rng3.End
      With rng3.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = myAltWord
        .Font.Underline = False
        .Wrap = wdFindContinue
        .MatchWholeWord = True
        .Replacement.Text = "^&!"
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
      End With
      DoEvents
      myAltCount = testDoc.Content.End - myTot
      WordBasic.EditUndo
      DoEvents
      Selection.End = 0
      DoEvents
      With rng3.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Text = "^&"
        .Replacement.Font.StrikeThrough = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
      End With
      DoEvents
      
      If InStr(res.Text, myAltWord & myLead) = 0 Then
        myLine = myAltWord & myLead & Trim(Str(myAltCount)) & vbCr
        res.InsertAfter myLine
        ActiveDocument.Paragraphs(2).range.Text = myLine
        ActiveDocument.Paragraphs(2).range.Font.StrikeThrough = True
      End If
      rng2.Start = myEnd
      With rng2.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findWord
        .Font.Underline = False
        .Font.StrikeThrough = False
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .Forward = True
        .MatchWildcards = False
        .Execute
      End With
    Loop
    If myAltWord > "" Then
      myLine = myAccWord & myLead & Trim(Str(myCount)) & vbCr
      res.InsertAfter myLine
      ActiveDocument.Paragraphs(1).range.Text = myLine
    End If
  End If
  rng.Start = myEnd
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "[" & allAccents & "]"
    .Font.Underline = False
    .Wrap = wdFindStop
    .Replacement.Text = ""
    .Forward = True
    .MatchWildcards = True
    .Execute
  End With
Loop

testDoc.Close SaveChanges:=False
resultDoc.Activate
Set rng = ActiveDocument.Content
rng.Sort SortOrder:=wdSortOrderAscending

Selection.HomeKey Unit:=wdStory
Selection.TypeText "Accent Use" & vbCr
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)
If doingSeveralMacros = False Then
  Beep
Else
  FUT.Activate
End If
End Sub

Sub CapitAlyse()
' Paul Beverley - Version 12.01.23
' Analyses capitalised words

ignoreHeadings = True
' ignoreHeadings = False

ignoreWords = ",After,All,Although,Also,An,And,As,At,By,For,From,If,In,It,"
ignoreWords = ignoreWords & "Of,On,Our,The,This,Those,There,These,They,Up,We,"

timeStart = Timer
showTime = True

myScreenOff = True

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

If myScreenOff = True Then
  Application.ScreenUpdating = False
  On Error GoTo ReportIt
End If

' Count case after hyphenation
myTot = ActiveDocument.Range.End
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "([A-Z][a-z]{1,}-[A-Z][a-z]{1,})"
  .Replacement.Text = "^&!"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
numCaps = ActiveDocument.Range.End - myTot
If numCaps > 0 Then WordBasic.EditUndo

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "([A-Z][a-z]{1,}-[a-z]{1,})"
  .Replacement.Text = "^&!"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
numNoCaps = ActiveDocument.Range.End - myTot
If numNoCaps > 0 Then WordBasic.EditUndo

Selection.HomeKey Unit:=wdStory
Set rng = ActiveDocument.Content
Documents.Add
Selection.FormattedText = rng.FormattedText
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ": "
  .Wrap = wdFindContinue
  .Replacement.Text = ". "
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = """"
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "[.]{2,}"
  .Replacement.Text = "."
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "(Figure [0-9]{1,}.[0-9]{1,})"
  .Replacement.Text = "\1. "
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "(Fig. [0-9]{1,}.[0-9]{1,})"
  .Replacement.Text = "\1. "
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "[^32]{1,}^13"
  .Replacement.Text = "^p"
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "^13[0-9.\)^t^32" & ChrW(8211) & "]{1,}([A-Z])"
  .Replacement.Text = "^p\1"
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "^13[a-z][.\)\(^t^32" & ChrW(8211) & "]{1,}([A-Z])"
  .Replacement.Text = "^p\1"
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "^t"
  .Wrap = wdFindContinue
  .Replacement.Text = ". "
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = ""
  .Wrap = wdFindContinue
  .Font.StrikeThrough = True
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
StatusBar = "Preparing the text for searching..."

' Underline headings to potentially ignore them
If ignoreHeadings = True Then
   maxWdsInHeading = 20
   For Each pa In ActiveDocument.Paragraphs
     myText = pa.Range.Text
     wdsCount = pa.Range.Words.count
     If Len(myText) > 3 And wdsCount < maxWdsInHeading Then
       ch = Mid(myText, Len(myText) - 1, 1)
       If InStr(".:", ch) = 0 Then pa.Range.Font.Underline = True
     End If
     i = i + 1: If i Mod 100 = 0 Then DoEvents
   Next pa
End If

For Each se In ActiveDocument.Sentences
  If Len(se) > 4 Then
    If InStr("""'(" & ChrW(8216) & ChrW(8220), _
         Trim(se.Words(1))) = 0 Then
      se.Words(1).Font.Underline = True
    Else
      se.Words(2).Font.Underline = True
    End If
    i = i + 1: If i Mod 500 = 0 Then DoEvents
  End If
Next se


With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "<[A-Z][a-zA-Z]{1,}"
  .Font.Underline = False
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute
End With

myCount = 0
myBars = "________________________________________"
allWords = "," & ignoreWords & ","
myResult = ""
Set tst = ActiveDocument.Content
myTot = tst.End
Do While rng.Find.Found = True
  endNow = rng.End
  If InStr(allWords, rng) = 0 Then
    StatusBar = myBars & myBars & myExtra & _
         "    >>> " & Int((myTot - endNow) / 1000)
    testWd = rng.Text
    allWords = allWords & testWd & ","
    lc = LCase(testWd)
    With tst.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = lc
      .MatchCase = True
      .Replacement.Text = "^&!"
      .MatchWildcards = False
      .MatchWholeWord = True
      .Execute Replace:=wdReplaceAll
    End With
    DoEvents
    numLC = ActiveDocument.Range.End - myTot
    If numLC > 0 Then
      WordBasic.EditUndo
      With tst.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = testWd
        .MatchCase = True
        .Replacement.Text = "^&!"
        .Execute Replace:=wdReplaceAll
      End With
      i = i + 1: If i Mod 20 = 0 Then DoEvents
      numCapAll = ActiveDocument.Range.End - myTot
      If numCapAll > 0 Then WordBasic.EditUndo
      With tst.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = testWd
        .Replacement.Text = "^&!"
        .Font.Underline = True
        .Execute Replace:=wdReplaceAll
      End With
      If i Mod 20 = 0 Then DoEvents
      numCapStart = ActiveDocument.Range.End - myTot
      numCapMid = numCapAll - numCapStart
      myExtra = lc & " . ." & Str(numLC) & "____:____"
      myExtra = myExtra & testWd & " . ." & Str(numCapMid)
      If numCapStart > 0 Then
        WordBasic.EditUndo
        myExtra = myExtra & " (+" & Trim(Str(numCapStart)) & ")"
      End If
      myResult = myResult & myExtra & ":" & vbCr
      If doingSeveralMacros = False Then _
           Debug.Print myExtra & "    >>> " & _
           Int((myTot - endNow) / 1000)
      myCount = myCount + 1
    End If
    rng.Start = endNow
    rng.End = endNow
  End If
  rng.Find.Execute
Loop

Selection.WholeStory
Selection.TypeText myResult
Selection.WholeStory
Selection.Range.Style = ActiveDocument.Styles(wdStyleNormal)
Selection.Font.Reset
Selection.Sort
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ":"
  .Replacement.Text = vbCr
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "_"
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With
Selection.HomeKey Unit:=wdStory
Selection.MoveEndWhile cset:=vbCr, count:=wdForward
Selection.Delete
Selection.TypeText "Capitalisation" & vbCr & vbCr
ActiveDocument.Paragraphs(1).Style = _
     ActiveDocument.Styles(wdStyleHeading1)
     Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "\(*\)"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Replacement.Font.Color = wdColorGray50
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

If numCaps + numNoCaps > 0 Then
  Selection.EndKey Unit:=wdStory
  myText = vbCr & vbCr & "Lowercase after hyphen (Non-linear): " & _
       Trim(Str(numNoCaps)) & vbCr & _
       "Uppercase after hyphen (Non-Linear): " & Trim(Str(numCaps)) _
       & vbCr & vbCr
  Selection.TypeText Text:=myText
  Selection.HomeKey Unit:=wdStory
End If

Application.ScreenUpdating = True
If doingSeveralMacros = False Then
  Beep
  myTime = Timer
  Do
  Loop Until Timer > myTime + 0.2
  Beep
  
  totTime = Timer - timeStart
  If showTime = True Then _
    MsgBox ((Int(10 * totTime / 60) / 10) & _
         "  minutes") & vbCr & vbCr & "Word pairs: " & myCount
Else
  FUT.Activate
End If
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub

Sub CenturyAlyse()
' Paul Beverley - Version 25.11.20
' Analyses how centuries are formatted in a document


Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

Set rng = ActiveDocument.Content
Documents.Add
Selection.FormattedText = rng.FormattedText

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "-"
  .Replacement.Text = " "
  .Wrap = wdFindContinue
  .MatchCase = False
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[sth]{2}"
  .Font.Superscript = True
  .Replacement.Text = "zcthzc"
  .Wrap = wdFindContinue
  .MatchCase = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totWas = ActiveDocument.Content.End
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "C[0-9]{2}>"
  .Replacement.Text = "11"
  .Replacement.Highlight = True
  .Wrap = wdFindContinue
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n01 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "C[0-9]{2}[ths]{2}>"
  .Replacement.Text = "2222"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n02 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "C[0-9]{2}zc[ths]{2}zc>"
  .Replacement.Text = "33333333"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n03 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[0-9]{2}zc[ths]{2}zc Ce"
  .Replacement.Text = "6666666666"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n06 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[0-9]{2}zc[ths]{2}zc ce"
  .Replacement.Text = "7777777777"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n07 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[XIV]{2}zc[ths]{2}zc Ce"
  .Replacement.Text = "8888888888"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n08 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[XIV]{2}zc[ths]{2}zc ce"
  .Replacement.Text = "9999999999"
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n09 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[XIV]{2}[ths]{2} Ce"
  .Replacement.Text = "AAAAAA"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n10 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[XIV]{2}[ths]{2} ce"
  .Replacement.Text = "BBBBBB"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n11 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[ienrlf]{2}[sth]{2} Ce"
  .Replacement.Text = "CCCCCC"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n12 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[ienrlf]{2}[sth]{2} ce"
  .Replacement.Text = "DDDDDD"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n13 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[ths]{2} Ce"
  .Replacement.Text = "4444"
  .Wrap = wdFindContinue
  .MatchCase = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n04 = totWas - totNow

totWas = totNow
With rng.Find
  .Text = "[ths]{2} ce"
  .Replacement.Text = "5555"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

totNow = ActiveDocument.Content.End
n05 = totWas - totNow

myReport = "C19YYn01QQC19thYYn02QQC19^thYYn03QQNineteenth " & _
     "CenturyYYn12QQnineteenth centuryYYn13QQ19th CenturyYY" & _
     "n04QQ19th centuryYYn05QQ19^th CenturyYYn06QQ19^th " & _
     "centuryYYn07QQXIXth CenturyYYn10QQXIXth centuryYYn11QQ" & _
     "XIX^th CenturyYYn08QQXIX^th centuryYYn09"

myReport = Replace(myReport, "YY", vbTab)
myReport = Replace(myReport, "QQ", vbCr)
myReport = Replace(myReport, "n01", Str(n01))
myReport = Replace(myReport, "n02", Str(n02))
myReport = Replace(myReport, "n03", Str(n03))
myReport = Replace(myReport, "n04", Str(n04))
myReport = Replace(myReport, "n05", Str(n05))
myReport = Replace(myReport, "n06", Str(n06))
myReport = Replace(myReport, "n07", Str(n07))
myReport = Replace(myReport, "n08", Str(n08))
myReport = Replace(myReport, "n09", Str(n09))
myReport = Replace(myReport, "n10", Str(n10))
myReport = Replace(myReport, "n11", Str(n11))
myReport = Replace(myReport, "n12", Str(n12))
myReport = Replace(myReport, "n13", Str(n13))

Selection.WholeStory
Selection.Range.Style = ActiveDocument.Styles(wdStyleNormal)
Selection.Font.Reset
Selection.TypeText myReport
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^^th"
  .Replacement.Text = "th"
  .Replacement.Font.Superscript = True
  .Wrap = wdFindContinue
  .MatchCase = True
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With
Selection.WholeStory
Selection.MoveEnd , -1
Selection.ConvertToTable Separator:=wdSeparateByTabs
Selection.Tables(1).Style = "Table Grid"
Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
Selection.Tables(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderVertical).LineStyle = wdLineStyleNone
Selection.Collapse wdCollapseStart
If doingSeveralMacros = False Then
  Beep
Else
  FUT.Activate
End If
End Sub

Sub DocAlyse()
' Paul Beverley - Version 21.12.23
' Analyses various aspects of a document

' prompts to count number of tests
cc = 53

myScreenOff = True

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)
If doingSeveralMacros = False Then
  myResponse = MsgBox("    DocAlyse" & vbCr & vbCr & _
       "Analyse this document?", vbQuestion _
       + vbYesNoCancel, "DocAlyse")
  If myResponse <> vbYes Then Exit Sub
End If

If myScreenOff = True Then
  Application.ScreenUpdating = False
  On Error GoTo ReportIt
End If

ss = "________________________________" & _
     "__________________________________"
myTrack = ActiveDocument.TrackRevisions
ActiveDocument.TrackRevisions = False

Set rngOld = ActiveDocument.Content
Documents.Add
Set tempDoc = ActiveDocument
Set rng = ActiveDocument.Content
rng.FormattedText = rngOld.FormattedText
Selection.EndKey Unit:=wdStory

If ActiveDocument.Footnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdFootnotesStory).Copy
  Selection.Paste
  Selection.Collapse wdCollapseEnd
End If

If ActiveDocument.Endnotes.Count > 0 Then
  ActiveDocument.StoryRanges(wdEndnotesStory).Copy
  Selection.Paste
  Selection.Collapse wdCollapseEnd
End If

Set endTempfile = ActiveDocument.Content
endTempfile.Collapse wdCollapseEnd

' collect text in all the textboxes (if any)
sh = ActiveDocument.Shapes.Count
If sh > 0 Then
  For Each shp In ActiveDocument.Shapes
    If shp.Type <> 24 And shp.Type <> 3 Then
      If shp.TextFrame.HasText Then
        shp.TextFrame.TextRange.Copy
        endTempfile.Select
        Selection.Paste
      End If
    End If
  Next
End If

Selection.HomeKey Unit:=wdStory


' Use main file for italic 'et al' count...
myTot = ActiveDocument.Range.End
Set rng = ActiveDocument.Content

cc = cc - 1
DoEvents

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "<et al>"
  .Font.Italic = True
  .Replacement.Text = "^&!"
  .Wrap = wdFindContinue
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute Replace:=wdReplaceAll
End With
italEtAls = ActiveDocument.Range.End - myTot
If italEtAls > 0 Then WordBasic.EditUndo

' ...and superscript degree count
cc = cc - 1
DoEvents
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[oO0]"
  .Font.Superscript = True
  .Replacement.Text = "vbvb"
  .Replacement.Font.Superscript = False
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
funnyDegrees = (ActiveDocument.Range.End - myTot) / 3

  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = " vbvb"
  .Replacement.Text = "^&!"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
funnyDegreesSp = ActiveDocument.Range.End - myTot - funnyDegrees * 3
If funnyDegreesSp > 0 Then WordBasic.EditUndo
If funnyDegrees > 0 Then WordBasic.EditUndo

DoEvents
Selection.HomeKey Unit:=wdStory
Set rngOld = ActiveDocument.Content
ActiveDocument.TrackRevisions = myTrack

Documents.Add
Set rng = ActiveDocument.Content
rng.FormattedText = rngOld.FormattedText
myEnd = rng.End
Set rng2 = ActiveDocument.Content
rng.Collapse wdCollapseEnd
rng.Text = rng2.Text

Set rng3 = ActiveDocument.Content
rng3.End = myEnd - 1
rng3.Select
Selection.Delete
myRslt = ""
Set rng = ActiveDocument.Content
myTot = ActiveDocument.Range.End
CR = vbCr: CR2 = CR & CR
tr = Chr(9) & "0zczc" & CR: SP = "     "
Selection.HomeKey Unit:=wdStory

Set newDoc = ActiveDocument

' Ten or 10
cc = cc - 1
DoEvents
myTot = ActiveDocument.Range.End
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "<ten>"
  .Replacement.Text = "!^&"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = " <10>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "ten" & vbTab & _
     Trim(Str(i)) & CR & "10" & vbTab & Trim(Str(g)) & CR2

' spelt-out lower-case numbers over nine
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<[efnst][efghinorvwx]{2,4}ty"
rng.Find.Execute Replace:=wdReplaceAll
aa = ActiveDocument.Range.End - myTot
If aa > 0 Then WordBasic.EditUndo

rng.Find.Text = "<ten>"
rng.Find.Execute Replace:=wdReplaceAll
ab = ActiveDocument.Range.End - myTot
If ab > 0 Then WordBasic.EditUndo

rng.Find.Text = "<eleven>"
rng.Find.Execute Replace:=wdReplaceAll
ac = ActiveDocument.Range.End - myTot
If ac > 0 Then WordBasic.EditUndo

rng.Find.Text = "<twelve>"
rng.Find.Execute Replace:=wdReplaceAll
ad = ActiveDocument.Range.End - myTot
If ad > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[efnst][efghinuorvwx]{2,4}teen>"
rng.Find.Execute Replace:=wdReplaceAll
ae = ActiveDocument.Range.End - myTot
If ae > 0 Then WordBasic.EditUndo

rng.Find.Text = "<hundred>"
rng.Find.Execute Replace:=wdReplaceAll
af = ActiveDocument.Range.End - myTot

If af > 0 Then WordBasic.EditUndo
If aa + ab + ac + ad + ae + af > 0 Then myRslt = myRslt & _
     "spelt-out numbers (11-999)" & vbTab & _
     Trim(Str(aa + ab + ac + ad + ae + af)) & CR2


' Four-digit numbers
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[!.]<[0-9]{4}>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

' take off 20xx dates
rng.Find.Text = "[!.]<20[0-9]{2}>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

' take off 13xx to 19xx dates
rng.Find.Text = "[!.]<1[3-9][0-9]{2}>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo
i = i - g - k
If i < 0 Then i = 0

' Four figs with comma
rng.Find.Text = "[!.]<[0-9],[0-9]{3}>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

' Four figs with hard or ordinary space
rng.Find.Text = "[!.]<[0-9][^0160^32][0-9]{3}>[!,]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo
If i + g + k > 0 Then
  myRslt = myRslt & "Four-digit numbers:" & CR _
  & "nnnn" & vbTab & Trim(Str(i)) & CR _
       & "n,nnn" & vbTab & Trim(Str(g)) & CR _
       & "n nnn" & vbTab & Trim(Str(k)) & CR2
End If



' Dates with 'mid' in front
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "mid [0-9]{4}"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "mid-[0-9]{4}"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "mid[0-9]{4}"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If i + g + k > 0 Then
  myRslt = myRslt & "mid 1900(s)" & vbTab _
       & Trim(Str(i)) & CR & "mid-1900(s)" & vbTab & _
       Trim(Str(g)) & CR & "mid1900(s)" & vbTab & _
       Trim(Str(k)) & CR2
End If

cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "mid [0-9]{2}[!0-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "mid-[0-9]{2}[!0-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "mid[0-9]{2}[!0-9]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If i + g + k > 0 Then
  myRslt = myRslt & "mid 90(s)" & vbTab _
       & Trim(Str(i)) & CR & "mid-90(s)" & vbTab & _
       Trim(Str(g)) & CR & "mid90(s)" & vbTab & _
       Trim(Str(k)) & CR2
End If


' Serial comma/not serial comma
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[a-zA-Z\-]{1,}, [a-zA-Z\-]{1,}, and "
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt & "serial comma" & vbTab & Trim(Str(i)) & CR

rng.Find.Text = "[a-zA-Z\-]{1,}, [a-zA-Z\-]{1,} and "
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt & "no serial comma" & vbTab & Trim(Str(i)) & CR2


' hard spaces
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "^s"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

' hard hyphens
rng.Find.Text = "^~"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
myRslt = myRslt & "hard spaces" & vbTab & Trim(Str(i)) _
     & CR & "hard hyphens" & vbTab & Trim(Str(g)) & CR2



' Single/double quotes
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = ChrW(8216)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
singleCurl = i
myRslt = myRslt & "curly open single quote" & vbTab & _
     Trim(Str(i)) & CR

rng.Find.Text = ChrW(8220)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt & "curly open double quote" & vbTab & _
     Trim(Str(i)) & CR

rng.Find.Text = Chr(39)
rng.Find.MatchWildcards = True
rng.Find.MatchCase = True
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt & "straight single quote" & vbTab & _
     Trim(Str(i)) & CR

rng.Find.Text = Chr(34)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt & "straight double quote" & vbTab & _
     Trim(Str(i)) & CR2


' etc(.)
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<etc[!.]"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "<etc."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<etc. [A-Z]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<etc.^13"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If h + i + g + k > 0 Then myRslt = myRslt & "etc" & _
     vbTab & Trim(Str(h)) & CR & "etc." & vbTab & _
     Trim(Str(i - g - k)) & CR2


' et al(.)
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<et al[!.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<et al."
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If g + i + italEtAls > 0 Then myRslt = myRslt & "et al." _
     & vbTab & Trim(Str(g)) & CR & "et al (italic, total)" & _
     vbTab & Trim(Str(italEtAls)) & CR & "et al (no dot)" & _
     vbTab & Trim(Str(i)) & CR2


' i.e./ie
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "i.e."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
myRslt = myRslt

rng.Find.Text = "<ie>"
rng.Find.MatchWildcards = True
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "ie" & vbTab & Trim(Str(g)) & CR _
     & "i.e." & vbTab & Trim(Str(i)) & CR2




' e.g./eg
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "e.g."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<eg>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "eg" & vbTab & Trim(Str(g)) & CR _
      & "e.g." & vbTab & Trim(Str(i)) & CR2




' Initials with surnames
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<[A-Z]. [A-Z]. [A-Z][a-z]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[A-Z][a-z]{2,}, [A-Z]. [A-Z]. "
rng.Find.Execute Replace:=wdReplaceAll
i2 = ActiveDocument.Range.End - myTot
If i2 > 0 Then WordBasic.EditUndo
aBit = "J. L. B. Matekoni" & vbTab & Trim(Str(i + i2)) & CR
g = i + i2

rng.Find.Text = "<[A-Z].[A-Z]. [A-Z][a-z]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[A-Z][a-z]{2,}, [A-Z].[A-Z]."
rng.Find.Execute Replace:=wdReplaceAll
i2 = ActiveDocument.Range.End - myTot
If i2 > 0 Then WordBasic.EditUndo
aBit = aBit & "J.L.B. Matekoni" & vbTab & Trim(Str(i + i2)) & CR
g = g + i + i2

rng.Find.Text = "<[A-Z] [A-Z] [A-Z][a-z]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[A-Z][a-z]{2,}, [A-Z] [A-Z] "
rng.Find.Execute Replace:=wdReplaceAll
i2 = ActiveDocument.Range.End - myTot
If i2 > 0 Then WordBasic.EditUndo
aBit = aBit & "J L B Matekoni" & vbTab & Trim(Str(i + i2)) & CR
g = g + i + i2

rng.Find.Text = "<[A-Z]{2}> [A-Z][a-z]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[A-Z][a-z]{2,}, [A-Z]{2}"
rng.Find.Execute Replace:=wdReplaceAll
i2 = ActiveDocument.Range.End - myTot
If i2 > 0 Then WordBasic.EditUndo
aBit = aBit & "JLB Matekoni" & vbTab & Trim(Str(i + i2)) & _
     "   (Beware! This can be inflated by, e.g. BBC Enterprises.)" & CR2




' Convention for page numbers
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<p. [1-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<pp. [1-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

k = i + g
aBit = "p/pp. 123" & vbTab & Trim(Str(k)) & CR

rng.Find.Text = "<p.[1-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<pp.[1-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

aBit = aBit & "p/pp.123" & vbTab & Trim(Str(i + g)) & CR
k = k + i + g

rng.Find.Text = "<p [1-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<pp [1-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

aBit = aBit & "p/pp 123" & vbTab & Trim(Str(i + g)) & CR
k = k + i + g

rng.Find.Text = "<p[1-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<pp[1-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

aBit = aBit & "p/pp123" & vbTab & Trim(Str(i + g)) & CR2
If k + i + g > 0 Then myRslt = myRslt & aBit




' Convention for ed/eds/edn
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<ed>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<eds>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<edn>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "<ed."
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = "<eds."
rng.Find.Execute Replace:=wdReplaceAll
m = ActiveDocument.Range.End - myTot
If m > 0 Then WordBasic.EditUndo

rng.Find.Text = "<edn."
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

If k + m + j + i + g + h > 0 Then myRslt = myRslt _
     & "ed" & vbTab & Trim(Str(i)) & CR & "eds" _
     & vbTab & Trim(Str(g)) & CR & "edn" & vbTab & _
       Trim(Str(h)) & CR & "ed." _
     & vbTab & Trim(Str(k)) & CR & "eds." & vbTab & _
       Trim(Str(m)) & CR & "edn." _
     & vbTab & Trim(Str(j)) & CR2



' Convention for am/pm
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[1-9][ap]m"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo
aBit = "2pm" & vbTab & Trim(Str(i)) & CR

rng.Find.Text = "[1-9][ap].m."
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
aBit = aBit & "2p.m." & vbTab & Trim(Str(g)) & CR

rng.Find.Text = "[1-9] [ap]m"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo
aBit = aBit & "2 pm" & vbTab & Trim(Str(k)) & CR

rng.Find.Text = "[1-9] [ap].m."
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo
aBit = aBit & "2 p.m." & vbTab & Trim(Str(h)) & CR2

If k + i + g + h > 0 Then myRslt = myRslt & aBit




' US/UK spelling
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[bpiv]our[ ,.s]"
rng.Find.Execute Replace:=wdReplaceAll
a = ActiveDocument.Range.End - myTot
If a > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,}elling>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,}elled>"
rng.Find.Execute Replace:=wdReplaceAll
f = ActiveDocument.Range.End - myTot
If f > 0 Then WordBasic.EditUndo


cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[bpiv]or[ ,.s]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "rior[ ,.s]"
rng.Find.Execute Replace:=wdReplaceAll
q = ActiveDocument.Range.End - myTot
If q > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,}eling>"
rng.Find.Execute Replace:=wdReplaceAll
v = ActiveDocument.Range.End - myTot
If v > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,}eled>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If a + g + f + i + q + v + k > 0 Then myRslt = _
     myRslt & "UK spelling (appx)" & vbTab & _
     Trim(Str(a + g + f)) & CR & _
     "US spelling (appx)" & vbTab & _
     Trim(Str(i - q + v + k)) & CR & _
     "(For a more accurate count, please use UKUScount.)" & CR2



' US/UK punctuation
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[a-zA-Z]['""" & ChrW(8217) & ChrW(8221) & "][,.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-zA-Z][,.]['""" & ChrW(8217) & ChrW(8221) & "][,.]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If i + k > 0 Then myRslt = myRslt & _
     "UK punctuation (appx)" & vbTab & _
     Trim(Str(i)) & CR & "US punctuation (appx)" _
     & vbTab & Trim(Str(k)) & CR2




' Initial capital after colon?
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[a-zA-Z]: [A-Z][a-z]"
rng.Find.Execute Replace:=wdReplaceAll
dfgsdfg = ActiveDocument.Range.End
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-zA-Z]: [a-z]"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

If i + j > 0 Then myRslt = myRslt & _
     "Initial capital after colon" & vbTab & _
     Trim(Str(i)) & CR & "Lowercase after colon" _
     & vbTab & Trim(Str(j)) & CR2



' is/iz
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "ise>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "ise[sd]>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "ising>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = "isation"
rng.Find.Execute Replace:=wdReplaceAll
l = ActiveDocument.Range.End - myTot
If l > 0 Then WordBasic.EditUndo

rng.Find.Text = "[armvt]ising"
rng.Find.Execute Replace:=wdReplaceAll
p = ActiveDocument.Range.End - myTot
If p > 0 Then WordBasic.EditUndo

rng.Find.Text = "[arvtw]ise"
rng.Find.Execute Replace:=wdReplaceAll
q = ActiveDocument.Range.End - myTot
If q > 0 Then WordBasic.EditUndo

rng.Find.Text = "ex[eo]rcis[ei]"
rng.Find.Execute Replace:=wdReplaceAll
r = ActiveDocument.Range.End - myTot
If r > 0 Then WordBasic.EditUndo
myRslt = myRslt & "-is- (very appx)" & vbTab & _
     Trim(Str(i + g + k + l - p - q - r)) & CR



cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "ize>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "ize[sd]>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "izing>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = "ization"
rng.Find.Execute Replace:=wdReplaceAll
l = ActiveDocument.Range.End - myTot
If l > 0 Then WordBasic.EditUndo

rng.Find.Text = "[Pp]riz[ie]"
rng.Find.Execute Replace:=wdReplaceAll
p = ActiveDocument.Range.End - myTot
If p > 0 Then WordBasic.EditUndo

rng.Find.Text = "[Sse]@iz[ie]"
rng.Find.Execute Replace:=wdReplaceAll
q = ActiveDocument.Range.End - myTot
If q > 0 Then WordBasic.EditUndo

myRslt = myRslt & "-iz- (very appx)" & vbTab _
     & Trim(Str(i + g + k + l - p - q)) & CR & _
     "(For a more accurate count, please use IZIScount.)" _
     & CR2




' data singular/plural
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<data is>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<data has>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<data was>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[Tt]his data>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo
myRslt = myRslt
l = i + g + h + k

cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
' If useVoice = True Then speech.Speak cc, SVSFPurgeBeforeSpeak
rng.Find.Text = "<data are>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<data have>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<data were>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "<[Tt]hese data>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo
If l + i + h + g + k > 0 Then myRslt = myRslt & _
     "data singular" & _
     vbTab & Trim(Str(l)) & CR & "data plural" & _
     vbTab & Trim(Str(i + g + h + k)) & CR2


' Is "first order" hyphenated?
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[snrt][tdh] order"
rng.Find.Execute Replace:=wdReplaceAll
dfgsdfg = ActiveDocument.Range.End
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[snrt][tdh]-order"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

If i + j > 0 Then myRslt = myRslt & _
     "'xxx order' with space" & vbTab & _
     Trim(Str(i)) & CR & "'xxx-order' with hyphen" _
     & vbTab & Trim(Str(j)) & CR2



' Is "three dimensional" hyphenated?
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[oweunvt ][eonrN1-4i] dimensional"
rng.Find.Execute Replace:=wdReplaceAll
dfgsdfg = ActiveDocument.Range.End
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[oweunvt ][eonrN1-4i]-dimensional"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

If i + j > 0 Then myRslt = myRslt & _
     "'xxx dimensional' with space" & vbTab & _
     Trim(Str(i)) & CR & "'xxx-dimensional' with hyphen" _
     & vbTab & Trim(Str(j)) & CR2




' Ellipsis, etc spacing
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr

allChars = "/" & ChrW(8211) _
     & ChrW(8212) & "-" & ChrW(8230)
myNames = "Solidus    En dash    Em dash    Hyphen     Ellipsis   Triple dotsSpaced dots  "
For myGo = 0 To 6
  sol = Mid(allChars, myGo + 1, 1)
  If myGo = 5 Then sol = "..."
  If myGo = 6 Then sol = ". . ."
  myName = Trim(Mid(myNames, (11 * myGo) + 1, 11))
  rng.Find.Text = sol
  rng.Find.Execute Replace:=wdReplaceAll
  t = ActiveDocument.Range.End - myTot
  If t > 0 Then
    WordBasic.EditUndo
    rng.Find.Text = " " & sol & " "
    rng.Find.Execute Replace:=wdReplaceAll
    bth = ActiveDocument.Range.End - myTot
    If bth > 0 Then WordBasic.EditUndo
    
    rng.Find.Text = "[! ]" & sol & " "
    rng.Find.MatchWildcards = True
    rng.Find.Execute Replace:=wdReplaceAll
    ftr = ActiveDocument.Range.End - myTot
    If ftr > 0 Then WordBasic.EditUndo
    
    rng.Find.Text = " " & sol & "[! ]"
    rng.Find.Execute Replace:=wdReplaceAll
    bfr = ActiveDocument.Range.End - myTot
    If bfr > 0 Then WordBasic.EditUndo
    
    rng.Find.Text = "[! ]" & sol & "[! ]"
    rng.Find.Execute Replace:=wdReplaceAll
    nthr = ActiveDocument.Range.End - myTot
    If nthr > 0 Then WordBasic.EditUndo
    
    myRslt = myRslt & myName & " spacing:" & CR & "space before only" _
         & vbTab & Trim(Str(bfr)) & CR & "space after only" & _
         vbTab & Trim(Str(ftr)) & CR & "spaced both ends" & _
         vbTab & Trim(Str(bth)) & CR
    If myGo <> 3 Then
      myRslt = myRslt & "not spaced" & vbTab & Trim(Str(nthr)) & CR2
    Else
      myRslt = myRslt & CR
    End If
    myRslt = myRslt & CR
  End If
Next myGo


' Types of ellipsis
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = ChrW(8230)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "..."
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = ". . ."
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

If i + j + k > 0 Then
  myRslt = myRslt & "Types of ellipsis:" & CR & _
       "proper ellipsis" & vbTab & Trim(Str(i)) & CR _
       & "triple dots" & vbTab & Trim(Str(j)) & CR _
       & "spaced triple dots" & vbTab & Trim(Str(k)) & CR2
End If



' line breaks
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "^l"
rng.Find.MatchWildcards = False
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

' page breaks
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "^m"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
myRslt = myRslt & "line breaks" & vbTab & Trim(Str(i)) _
  & CR & "page breaks" & vbTab & Trim(Str(g)) & CR2



' fig/figure
aBit = ""
rng.Find.Text = "<fig>[!.]"
rng.Find.MatchWildcards = True
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "fig" & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<Fig>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "Fig" & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<fig."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "fig." & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<Fig."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "Fig." & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<figs>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "figs" & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<Figs>[!.]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "Figs" & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "<figs."
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "figs." & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "figure [0-9\(]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "figure" & vbTab & Trim(Str(i)) & CR
End If

rng.Find.Text = "[!.] Figure [0-9\(]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then
  WordBasic.EditUndo
  aBit = aBit & "Figure" & vbTab & Trim(Str(i)) & CR
End If
If aBit > "" Then myRslt = myRslt & aBit & CR




' Chapter/chapter
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[!.] Chapter [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "chapter [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then
  myRslt = myRslt & "Chapter (number)" & vbTab & Trim(Str(i)) & CR _
       & "chapter (number)" & vbTab & Trim(Str(g)) & CR2
End If


' Section/section
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[!.] Section [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "section [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then
  myRslt = myRslt & "Section (number)" & vbTab & _
       Trim(Str(i)) & CR & "section (number)" _
       & vbTab & Trim(Str(g)) & CR2
End If


' No./no.
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
' If useVoice = True Then speech.Speak cc, SVSFPurgeBeforeSpeak
rng.Find.Text = " No. [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = " No [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = " no. [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = " No.[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = " No[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
l = ActiveDocument.Range.End - myTot
If l > 0 Then WordBasic.EditUndo

rng.Find.Text = " no.[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
m = ActiveDocument.Range.End - myTot
If m > 0 Then WordBasic.EditUndo

If i + j + g + k + l + m > 0 Then
  myRslt = myRslt & "No (number)" & vbTab & Trim(Str(i)) _
     & CR & "No. (number)" & vbTab & Trim(Str(j)) & CR _
     & "no. (number)" & vbTab & Trim(Str(g)) & CR
  myRslt = myRslt & "No(number)" & vbTab & Trim(Str(k)) _
     & CR & "No.(number)" & vbTab & Trim(Str(l)) & CR _
     & "no.(number)" & vbTab & Trim(Str(m)) & CR2
End If

cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = " Vol. [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = " Vol [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = " vol. [0-9]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = " Vol.[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = " Vol[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
l = ActiveDocument.Range.End - myTot
If l > 0 Then WordBasic.EditUndo

rng.Find.Text = " vol.[0-9]"
rng.Find.Execute Replace:=wdReplaceAll
m = ActiveDocument.Range.End - myTot
If m > 0 Then WordBasic.EditUndo

If i + j + g + k + l + m > 0 Then
  myRslt = myRslt & "Vol (number)" & vbTab & Trim(Str(i)) _
      & CR & "Vol. (number)" & vbTab & Trim(Str(j)) & CR _
     & "vol. (number)" & vbTab & Trim(Str(g)) & CR
  myRslt = myRslt & "Vol(number)" & vbTab & Trim(Str(k)) _
     & CR & "Vol.(number)" & vbTab & Trim(Str(l)) & CR _
     & "vol.(number)" & vbTab & Trim(Str(m)) & CR
  myRslt = myRslt & CR
End If



' equations
fText = ",<eq [0-9],<eq. [0-9],<eqn [0-9],<Eqn [0-9]," & _
"eqns [0-9],eqs [0-9],<eq \(,<eq. \(,<Eq. \(" & _
",<eqn \(,<Eqn \(,eqns \(,eqs \(,Eqs \(,Eqs. \(" & _
",equation \(,[!.] Equation \(,equations \(" & _
",[!.] Equations \(,equation [0-9]" & _
",[!.] Equation [0-9],equations [0-9]" & _
",[!.] Equations [0-9]"

resText = ",eq,eq.,eqn,Eqn,eqns,eqs,eq (n.n)," & _
"eq. (n.n),Eq. (n.n),eqn (n.n)," & _
"Eqn (n.n),eqns (n.n),eqs,Eqs,Eqs.," & _
"equation (n.n),Equation (n.n)," & _
"equations (n.n),Equations (n.n),equation," & _
"Equation,equations,Equations"

myF = Split(fText, ",")
myRes = Split(resText, ",")

aBit = ""
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.MatchWildcards = True
For q = 1 To 23
  With rng.Find
    .Text = myF(q)
    .Replacement.Text = "^&!"
    .Execute Replace:=wdReplaceAll
  End With
  i = ActiveDocument.Range.End - myTot
  Debug.Print ActiveDocument.Range.End, myTot, i
  If i > 0 Then WordBasic.EditUndo: _
       aBit = aBit & myRes(q) & vbTab & Trim(Str(i)) & CR
Next q
If aBit > "" Then myRslt = myRslt & aBit & CR



' units
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[0-9][^32^160][kKcmM][NgAVm]>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9][^32^160][NgAVm]>"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9][kKcmM][NgAVm]>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9][NgAVm]>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo
If i + j + g + h > 0 Then
  myRslt = myRslt & "spaced units (3 mm)" & vbTab & _
       Trim(Str(i + j)) & CR & "unspaced units (3mm)" _
     & vbTab & Trim(Str(g + h)) & CR2
End If


' Ok, OK, ok, okay
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "<OK>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<Ok>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<ok>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "<okay>"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

If i + h + g + j > 0 Then myRslt = myRslt & "OK" & _
     vbTab & Trim(Str(i)) & CR _
     & "Ok" & vbTab & Trim(Str(g)) & CR _
     & "ok" & vbTab & Trim(Str(h)) & CR _
     & "okay" & vbTab & Trim(Str(j)) & CR2

' Now go to all lowercase
rng.Case = wdLowerCase
myTot = ActiveDocument.Range.End


' Backward(s), forward(s) etc.
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "[abcdfiknoprtuw]{2,4}ward>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[abcdfiknoprtuw]{2,4}wards>"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo


If i + j > 0 Then myRslt = myRslt & "back/for/toward etc." & _
     vbTab & Trim(Str(i)) & CR _
     & "back/for/towards etc." & vbTab & Trim(Str(j)) & CR2



' amid(st), among(st), while(st)
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "<amid>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "<among>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo: g = g + h

rng.Find.Text = "<while>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo: g = g + h


rng.Find.Text = "<amidst>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "<amongst>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo: i = i + h

rng.Find.Text = "<whilst>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo: i = i + h

If i + g > 0 Then
  myRslt = myRslt & "amid/among/while" & vbTab & Trim(Str(g)) & CR
  myRslt = myRslt & "amidst/amongst/whilst" & vbTab & Trim(Str(i)) & CR2
End If



' past participle -rnt -elt
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "sp[oi]@lt>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "lea[np]t>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "[l ][be][ua]rnt>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[ds][wpm]elt>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = "sp[oi]@[l]@ed>"
rng.Find.Execute Replace:=wdReplaceAll
p = ActiveDocument.Range.End - myTot
If p > 0 Then WordBasic.EditUndo

rng.Find.Text = "lea[np]ed>"
rng.Find.Execute Replace:=wdReplaceAll
q = ActiveDocument.Range.End - myTot
If q > 0 Then WordBasic.EditUndo

rng.Find.Text = "[l ][be][ua]rned>"
rng.Find.Execute Replace:=wdReplaceAll
r = ActiveDocument.Range.End - myTot
If r > 0 Then WordBasic.EditUndo

rng.Find.Text = "[ds][wpm]elled>"
rng.Find.Execute Replace:=wdReplaceAll
s = ActiveDocument.Range.End - myTot
If g + h + i + k + p + q + r + s > 0 Then myRslt = myRslt & _
     "-rnt -elt" & vbTab & Trim(Str(g + h + i + k)) & CR & _
     "-rned -elled" & vbTab & Trim(Str(p + q + r + s)) & CR2



' percentages
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "[0-9]%"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9][^32^160]%"
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9] per cent>"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9] percent>"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,} per cent>"
rng.Find.Execute Replace:=wdReplaceAll
k = ActiveDocument.Range.End - myTot
If k > 0 Then WordBasic.EditUndo

rng.Find.Text = "[a-z]{3,} percent>"
rng.Find.Execute Replace:=wdReplaceAll
m = ActiveDocument.Range.End - myTot
If m > 0 Then WordBasic.EditUndo

If i + j + g + h + k + m > 0 Then
  myRslt = myRslt & "unspaced, e.g.   9%" & vbTab & _
       Trim(Str(i)) & CR & "spaced, e.g.   9 %" _
     & vbTab & Trim(Str(j)) & CR & "9 per cent" & vbTab & _
       Trim(Str(g)) & CR & "9 percent" _
     & vbTab & Trim(Str(h)) & CR & "nine per cent" & vbTab & _
       Trim(Str(k)) & CR & "nine percent" _
     & vbTab & Trim(Str(m)) & CR2
End If



' Feet and inches
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
curlyOpt = Options.AutoFormatAsYouTypeReplaceQuotes
Options.AutoFormatAsYouTypeReplaceQuotes = False
rng.Find.Text = "[0-9]'"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9]"""
rng.Find.Execute Replace:=wdReplaceAll
j = ActiveDocument.Range.End - myTot
If j > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9]" & ChrW(8242)
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "[0-9]" & ChrW(8243)
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo
Options.AutoFormatAsYouTypeReplaceQuotes = curlyOpt

If i + j + g + h > 0 Then
  myRslt = myRslt & "feet (straight)   9'" & vbTab & _
       Trim(Str(i)) & CR & "inches (straight)   9""" _
       & vbTab & Trim(Str(j)) & CR & "single prime   9" & _
       ChrW(8242) & vbTab & Trim(Str(g)) & CR & _
       "double prime   9" & ChrW(8243) & vbTab & _
       Trim(Str(h)) & CR2
End If


' focus(s)
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "focus[ei]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "focuss[ei]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "focus..." & _
     vbTab & Trim(Str(i)) & CR _
     & "focuss..." & vbTab & Trim(Str(g)) & CR2



' benefit(t)
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "benefit[ei]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "benefitt[ei]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "benefit..." & _
     vbTab & Trim(Str(i)) & CR _
     & "benefitt..." & vbTab & Trim(Str(g)) & CR2



' co(-)oper...
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "co-op[ei]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "coop[ei]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "co-oper..." & _
     vbTab & Trim(Str(i)) & CR _
     & "cooper..." & vbTab & Trim(Str(g)) & CR2



' Co-ordin
rng.Find.Text = "co-ord[ei]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "coord[ei]"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo
If i + g > 0 Then myRslt = myRslt & "co-ord..." & _
     vbTab & Trim(Str(i)) & CR _
     & "coord..." & vbTab & Trim(Str(g)) & CR2



' Can't, cannot can not
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "can[!a-z]t>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "cannot"
rng.Find.Execute Replace:=wdReplaceAll
g = ActiveDocument.Range.End - myTot
If g > 0 Then WordBasic.EditUndo

rng.Find.Text = "can not"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

If i + h + g > 0 Then myRslt = myRslt & "can't" & _
     vbTab & Trim(Str(i)) & CR _
     & "cannot" & vbTab & Trim(Str(g)) & CR _
     & "can not" & vbTab & Trim(Str(h)) & CR2



' Wasn't, isn't, hasn't
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
myTot = ActiveDocument.Range.End
rng.Find.Text = "[owh ][aie]sn[!a-z]t>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[owh ][aie]s not"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

If i + h > 0 Then myRslt = myRslt & _
     "wasn't, isn't, hasn't" _
     & vbTab & Trim(Str(i)) & CR _
     & "was not, is not, has not" & vbTab & _
     Trim(Str(h)) & CR2



' Don't, won't
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "[dw]on[!a-z]t>"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "[dw][oil]{1,3} not"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

If i + h > 0 Then myRslt = myRslt & _
     "don't, won't" _
     & vbTab & Trim(Str(i)) & CR _
     & "do not, will not" & vbTab & _
     Trim(Str(h)) & CR2



' which/that
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = "which"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = "that"
rng.Find.Execute Replace:=wdReplaceAll
h = ActiveDocument.Range.End - myTot
If h > 0 Then WordBasic.EditUndo

If i + h > 0 Then myRslt = myRslt & _
     "which" _
     & vbTab & Trim(Str(i)) & CR _
     & "that" & vbTab & _
     Trim(Str(h)) & CR2



' Funny characters
cc = cc - 1
DoEvents
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
rng.Find.Text = ChrW(178)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo: _
     myRslt = myRslt & "funny 'squared' character" _
       & vbTab & Trim(Str(i)) & CR2

myDiacritics = ""
For i = 192 To 255
  If i <> 215 And i <> 247 Then
    myDiacritics = myDiacritics & ChrW(i)
  End If
Next i
rng.Find.Text = "[" & myDiacritics & "]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo: _
     myRslt = myRslt & "diacritics" & vbTab & Trim(Str(i)) & CR2

rng.Find.Text = "[" & ChrW(191) & ChrW(161) & ChrW(139) & _
     ChrW(155) & ChrW(171) & ChrW(187) & "]"
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo: myRslt = myRslt & _
     "Continental punctuation" & vbTab & Trim(Str(i)) & CR2

' Ordinary degree symbols
rng.Find.Text = ChrW(176)
rng.Find.Execute Replace:=wdReplaceAll

i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = " " & ChrW(176)
rng.Find.Execute Replace:=wdReplaceAll
isp = ActiveDocument.Range.End - myTot
If isp > 0 Then WordBasic.EditUndo
If i > 0 Then myRslt = myRslt & "degree symbols closed" _
      & vbTab & Trim(Str(i - isp)) & CR _
      & "degree symbols spaced" _
      & vbTab & Trim(Str(isp)) & CR2

' Funny degrees
rng.Find.Text = ChrW(186)
rng.Find.Execute Replace:=wdReplaceAll
i = ActiveDocument.Range.End - myTot
If i > 0 Then WordBasic.EditUndo

rng.Find.Text = " " & ChrW(186)
rng.Find.Execute Replace:=wdReplaceAll
isp = ActiveDocument.Range.End - myTot
If isp > 0 Then WordBasic.EditUndo

If i + funnyDegrees > 0 Then
  myRslt = myRslt & "funny degrees (0186) closed" _
      & vbTab & Trim(Str(i + funnyDegrees - isp - _
      funnyDegreesSp)) & CR _
      & "funny degrees (0186) spaced" _
      & vbTab & Trim(Str(isp + funnyDegreesSp)) & CR2
End If


appx = ""
If colouredText > 0 Then
  If colourOverflow = True Then appx = " (I think)"
  myRslt = myRslt & "text in coloured font" _
      & appx & vbTab & Trim(Str(colouredText - 1)) & CR2
End If

If lineBreaks > 0 Then
  myRslt = myRslt & "line breaks" _
      & vbTab & Trim(Str(i + lineBreaks)) & CR2
End If

If pageBreaks > 0 Then
  myRslt = myRslt & "page breaks" _
      & vbTab & Trim(Str(i + pageBreaks)) & CR2
End If


myRslt = myRslt & CR

Selection.HomeKey Unit:=wdStory
Selection.TypeText Text:="DocAlyse" & vbCr & vbCr
newDoc.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)

newDoc.Paragraphs(3).Range.Select
Selection.End = newDoc.Content.End
Selection.TypeText CR & myRslt & CR2
Selection.Font.Bold = True
Set rng = ActiveDocument.Content
rng.ParagraphFormat.TabStops.ClearAll
rng.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(4.5), _
    Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces

' Grey out the zero lines
cc = cc - 1
DoEvents
Set rng = ActiveDocument.Content
StatusBar = ss & "  " & Trim(Str(cc)) & vbCr
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^13([!^13]@)^t0"
  .Wrap = wdFindContinue
  .Replacement.Text = "^p\1^t^="
  .Replacement.Font.Bold = False
  .Replacement.Font.Color = wdColorGray25
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  
  .Text = "^t^=zczc"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

With Selection.Find
  .Text = ""
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute
End With

Selection.HomeKey Unit:=wdStory
tempDoc.Close SaveChanges:=False

If doingSeveralMacros = False Then
  Beep
  newDoc.Activate
Else
  FUT.Activate
End If
Application.ScreenUpdating = True
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub

Sub FullNameAlyse()
' Paul Beverley - Version 24.02.21
' Creates a frequency list of all full names

IncludeNamesWithInitials = vbYes

' In this list, make sure every word has a space after it
allowAbbrevs = "Mr. Mrs. Dr."

nonoWords = "About After Although An And Any As At Before Because " & _
     "But By For Has Have However If In Is Like My Since So Some " & _
     "That The Then These This Those Though Through Unlike " & _
     "Was We What When While Who Why Yet "

nonoWords2 = "an and are do no nor on one or v "


Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

myResponse = IncludeNamesWithInitials
If doingSeveralMacros = False Then
  myResponse = MsgBox("Include names with initials?", vbQuestion _
          + vbYesNoCancel, "FullNameAlyse")
  If myResponse = vbCancel Then Exit Sub
End If

Set rng = ActiveDocument.Content
Documents.Add
Set originalDoc = ActiveDocument
Selection.FormattedText = rng.FormattedText

' Now prepare the text
numberCmnts = ActiveDocument.Comments.Count
If numberCmnts > 0 Then ActiveDocument.DeleteAllComments

Set rng = ActiveDocument.Content
myEnd = rng.End
' Make apostrophes straight
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ChrW(8217)
  .Wrap = wdFindContinue
  .Replacement.Text = "'"
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "'s"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

thisArray = Split(Trim(allowAbbrevs), " ")
For i = 0 To UBound(thisArray)
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = thisArray(i)
    .Wrap = wdFindContinue
    .Replacement.Text = Replace(thisArray(i), ".", "")
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With
  DoEvents
Next i

Documents.Add
CR = vbCr

' First mark all two-word proper nouns, in order
' to detect four-word names (= two + two)
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = "<[A-Z][a-zA-Z\-']{1,} [A-Z][a-zA-Z\-']{1,}?"
 .Font.StrikeThrough = False
 .Wrap = wdFindStop
 .Replacement.Font.DoubleStrikeThrough = True
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute Replace:=wdReplaceAll
End With

' Find four-word names
rng.Start = 0
rng.End = 0
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = ""
 .Font.StrikeThrough = False
 .Font.DoubleStrikeThrough = True
 .Wrap = wdFindStop
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute
End With

Set firstDoc = ActiveDocument
Do While rng.Find.Found = True
numWords = rng.Words.Count
  If numWords > 2 And numWords < 7 Then
    myText = Left(rng.Text, Len(rng.Text) - 1)
    Selection.TypeText myText & CR
    rng.Font.Shadow = True
  End If
  rng.Collapse wdCollapseEnd
  rng.Find.Execute
  DoEvents
Loop


' Find three-word names
rng.Start = 0
rng.End = 0
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = "<[A-Z][a-zA-Z\-']{1,} [A-Z][a-zA-Z\-']{1,} [A-Z][a-zA-Z\-']{1,}"
 .Font.StrikeThrough = False
 .Font.Shadow = False
 .Wrap = wdFindStop
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute
End With

CR = vbCr
Set firstDoc = ActiveDocument
Do While rng.Find.Found = True
  Selection.TypeText rng.Text & CR
  rng.Font.Shadow = True
  rng.Collapse wdCollapseEnd
  rng.Find.Execute
  DoEvents
Loop

' Find three-word names with van, von, der, de etc
rng.Start = 0
rng.End = 0
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = "<[A-Z][a-zA-Z]{1,} [A-Z][a-zA-Z]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}"
 .Font.StrikeThrough = False
 .Font.Shadow = False
 .Wrap = wdFindStop
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute
End With

CR = vbCr
Set firstDoc = ActiveDocument
Do While rng.Find.Found = True
  Selection.TypeText rng.Text & CR
  rng.Font.Shadow = True
  rng.Collapse wdCollapseEnd
  rng.Find.Execute
  DoEvents
Loop

rng.Start = 0
rng.End = 0
' Two-word names
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = "<[A-Z][a-zA-Z\-']{1,} [A-Z][a-zA-Z\-']{1,}"
 .Font.StrikeThrough = False
 .Font.Shadow = False
 .Wrap = wdFindStop
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute
End With

Do While rng.Find.Found = True
  Selection.TypeText rng.Text & CR
  rng.Collapse wdCollapseEnd
  rng.Find.Execute
  DoEvents
Loop

rng.Start = 0
rng.End = 0
' Two-word names with van, von, der, de etc
With rng.Find
 .ClearFormatting
 .Replacement.ClearFormatting
 .Text = "<[A-Z][a-zA-Z]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}>"
 .Font.StrikeThrough = False
 .Font.Shadow = False
 .Wrap = wdFindStop
 .Replacement.Text = ""
 .Forward = True
 .MatchWildcards = True
 .Execute
End With

Do While rng.Find.Found = True
  Selection.TypeText rng.Text & CR
  rng.Collapse wdCollapseEnd
  rng.Find.Execute
  DoEvents
Loop

If myResponse = vbYes Then
  ' Find such as P.E. Beverley
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z.]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop
  
  ' Find such as Paul E. Beverley
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z][a-zA-Z]{1,} [A-Z.]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop

  ' Find such as P E Beverley
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z ]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop

  ' Find such as Paul E H Beverley
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z][a-zA-Z]{1,} [A-Z ]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop
  
  ' Find such as P.E. Beverley + van der etc
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z.]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop
  
  ' Find such as Paul E. Beverley + van der etc
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z][a-zA-Z]{1,} [A-Z.]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop
  
  ' Find such as P E Beverley + van der etc
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z ]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop

  ' Find such as Paul E H Beverley + van der etc
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z][a-zA-Z]{1,} [A-Z ]{1,} [vanderol]{1,} [A-Z][a-zA-Z]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    Selection.TypeText rng.Text & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop

  ' Find such as Beverley, P.E.
  rng.Start = 0
  rng.End = 0
  With rng.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = "<[A-Z][a-zA-Z]{1,}, [A-Z. ]{1,}>"
   .Font.StrikeThrough = False
   .Wrap = wdFindStop
   .Replacement.Text = ""
   .Forward = True
   .MatchWildcards = True
   .Execute
  End With
  
  Do While rng.Find.Found = True
    nameInits = rng.Text
    commaPos = InStr(nameInits, ",")
    initsName = Mid(nameInits, commaPos - 1) & " " & Left(nameInits, commaPos - 1)
    Selection.TypeText initsName & CR
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    DoEvents
  Loop
End If

rng.Start = 0
rng.End = myEnd
rng.Font.Shadow = False
rng.Font.DoubleStrikeThrough = False

Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
Selection.EndKey Unit:=wdStory
Selection.TypeText CR
Selection.HomeKey Unit:=wdStory
Selection.MoveEnd , 1
Selection.Delete

Dim myName(8000) As String
Dim itemCount As Long
Dim myCount As Integer
Dim thisPara As String
Dim prevPara As String

myCount = 0
prevName = ""
For Each myPara In ActiveDocument.Paragraphs
  thisPara = Replace(myPara.Range.Text, CR, "")
  If thisPara <> prevPara And prevPara <> "" Then
    itemCount = itemCount + 1
    myName(itemCount) = prevPara & vbTab & Trim(Str(myCount))
    myCount = 1
  Else
    myCount = myCount + 1
  End If
  prevPara = thisPara
  DoEvents
Next myPara

Documents.Add
Set secondDoc = ActiveDocument

For i = 1 To itemCount
  If UCase(myName(i)) <> myName(i) Then
    Selection.TypeText myName(i) & CR
  End If
  DoEvents
Next i

maxLine = ActiveDocument.Paragraphs.Count
nonoWords = nonoWords & " "
For i = maxLine To 1 Step -1
  firstWord = ActiveDocument.Paragraphs(i).Range.Words(1)
  DeleteIt = (InStr(nonoWords, firstWord) > 0)
  For j = 2 To ActiveDocument.Paragraphs(i).Range.Words.Count - 1
    thisWord = Trim(ActiveDocument.Paragraphs(i).Range.Words(j))
    If InStr(nonoWords2, thisWord & " ") > 0 Then DeleteIt = True
  Next j
  If DeleteIt = True Then ActiveDocument.Paragraphs(i).Range.Delete
Next i
totalItems = ActiveDocument.Paragraphs.Count - 1

' Copy the list and paste into the first document
' as a place to manipulate it
Selection.WholeStory
Selection.Copy
firstDoc.Activate
Selection.WholeStory
Selection.Delete
Selection.Paste

' Move the surname to the beginning of the line
For Each myPara In ActiveDocument.Paragraphs
  If Len(myPara.Range.Text) > 2 Then
    surnamePosn = myPara.Range.Words.Count - 3
    If InStr(myPara.Range.Text, "-") = 0 Then
      Surname = Trim(myPara.Range.Words(surnamePosn))
      myPara.Range.Words(surnamePosn) = ""
      myPara.Range.Words(1) = Surname & ", " & myPara.Range.Words(1)
    Else
      myPara.Range.Words(surnamePosn).Select
      Selection.MoveStartUntil cset:=" ", Count:=wdBackward
      Selection.MoveStart , -1
      fullSurname = Trim(Selection.Text)
      Selection.Delete
      Selection.HomeKey Unit:=wdLine
      Selection.TypeText fullSurname & ", "
      asdgfdfg = 0
    End If
  End If
  DoEvents
Next myPara

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = " ^t"
  .Replacement.Text = "^t"
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

' Format the list
Selection.HomeKey Unit:=wdStory
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric

Selection.HomeKey Unit:=wdStory
Selection.MoveEnd , 2
Selection.Delete
Selection.TypeText "Fullname List" & vbCr & vbCr
Selection.TypeText "Sorted by last name" & vbCr
startTable = Selection.End
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading2)
ActiveDocument.Paragraphs(3).Style = ActiveDocument.Styles(wdStyleHeading2)
Selection.Start = startTable
Selection.End = ActiveDocument.Range.End
Selection.ConvertToTable Separator:=wdSeparateByTabs
Selection.Tables(1).Style = "Table Grid"
Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
Selection.Tables(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderVertical).LineStyle = wdLineStyleNone

Selection.WholeStory
Selection.Copy
ActiveDocument.Close SaveChanges:=False

' Format other list
secondDoc.Activate
Selection.HomeKey Unit:=wdStory
Selection.TypeText "Sorted by first name" & vbCr
startTable = Selection.End
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading2)
Selection.Start = startTable
Selection.End = ActiveDocument.Range.End
Selection.ConvertToTable Separator:=wdSeparateByTabs
Selection.Tables(1).Style = "Table Grid"
Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
Selection.Tables(1).Borders(wdBorderTop).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderRight).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
Selection.Tables(1).Borders(wdBorderVertical).LineStyle = wdLineStyleNone

' Copy the other list in here
Selection.HomeKey Unit:=wdStory
Selection.Paste
Selection.HomeKey Unit:=wdStory

' Dummy copy to clear clipboard
Set rng = ActiveDocument.Content
rng.End = rng.Start + 1
rng.Copy
originalDoc.Activate
ActiveDocument.Close SaveChanges:=False

If doingSeveralMacros = False Then
  Beep
  MsgBox (Str(totalItems) & " names found")
Else
  FUT.Activate
End If
End Sub

Sub HyphenAlyse()
' Paul Beverley - Version 04.04.22
' Creates a frequency list of all possible hyphenations

myList = "anti,cross,eigen,hyper,inter,meta,mid,multi," _
     & "non,over,post,pre,pseudo,quasi,semi,sub,super"
    
includeNumbers = True

deleteTableBorders = True
lighterColour = wdGray25
' lighterColour = wdColor50

Dim myResult As String
On Error GoTo ReportIt
myList = "," & myList
myList = Replace(myList, ",,", ",")
pref = Split(myList, ",")

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

If doingSeveralMacros = False Then
  myResponse = MsgBox("    HyphenAlyse" & vbCr & vbCr & _
       "Analyse hyphenated words?", vbQuestion _
       + vbYesNoCancel, "HyphenAlyse")
  If myResponse <> vbYes Then Exit Sub
End If
Dim pr(8000) As String
Set mainDoc = ActiveDocument
strttime = Timer
Set rng = ActiveDocument.Content
Documents.Add
Selection.FormattedText = rng.FormattedText
Selection.HomeKey Unit:=wdStory
Application.ScreenUpdating = False
On Error GoTo ReportIt

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Font.StrikeThrough = True
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Set rng = ActiveDocument.Content
rng.Case = wdLowerCase
Set tempDoc = ActiveDocument
Documents.Add
Set newTemp = ActiveDocument
Selection.Text = rng.Text
tempDoc.Close SaveChanges:=False
newTemp.Activate

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ChrW(8217) & "[!a-z]"
  .Wrap = wdFindContinue
  .Replacement.Text = "!!"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With

If includeNumbers = True Then
  schStr = "[a-z0-9]{1,}[-^=][0-9a-z-]{1,}"
Else
  schStr = "[a-z]{1,}[-^=][a-z-]{1,}"
End If
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = schStr
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute
End With

' Find all hyphenated/dashed word pairs
myPairs = 0
allWords = ","
Do While rng.Find.Found = True
  wdPair = Replace(rng.Text, ChrW(8211), "-")
  If InStr(allWords, "," & wdPair & ",") = 0 _
       And (UCase(wdPair) <> wdPair) Then
    myPairs = myPairs + 1
    pr(myPairs) = wdPair
    allWords = allWords & wdPair & ","
    If myPairs Mod 10 = 0 Then
      If doingSeveralMacros = False Then _
           Debug.Print rng.Text, myPairs
      StatusBar = rng.Text & "     " & myPairs
    End If
  End If
  If Right(wdPair, 1) <> "s" Then
    wdPairs = wdPair & "s"
    If InStr(allWords, "," & wdPairs & ",") = 0 Then
      myPairs = myPairs + 1
      pr(myPairs) = wdPairs
      allWords = allWords & wdPairs & ","
      If myPairs Mod 10 = 0 Then
        If doingSeveralMacros = False Then _
             Debug.Print rng.Text, myPairs
        StatusBar = rng.Text & "     " & myPairs
      End If
    End If
  End If
  rng.Find.Execute
  DoEvents
Loop

' Collect words with each prefix
For i = 1 To UBound(pref)
  hPos = Len(pref(i))
  allPreWords = ","
  
  If includeNumbers = True Then
    schStr = "<" & pref(i) & "[0-9a-z]{2,}"
  Else
    schStr = "<" & pref(i) & "[a-z]{2,}"
  End If
  
  Set rng = ActiveDocument.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = schStr
    .Wrap = wdFindStop
    .Replacement.Text = ""
    .Forward = True
    .MatchWildcards = True
    .Execute
  End With
  
  Do While rng.Find.Found = True
    wd = rng.Text
    If InStr(wd, "-") = 0 Then wd = Left(wd, hPos) _
         & "-" & Mid(wd, hPos + 1)
    If InStr(allPreWords, "," & wd & ",") = 0 And _
         InStr(allWords, "," & wd & ",") = 0 Then
      myPairs = myPairs + 1
      pr(myPairs) = wd
      allPreWords = allPreWords & wd & ","
      allWords = allWords & wd & ","
      If myPairs Mod 10 = 0 Then
        If doingSeveralMacros = False Then _
             Debug.Print wd, myPairs
        StatusBar = wd & "     " & myPairs
      End If
    End If
    DoEvents
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
  Loop
Next i

' Collect word pairs with each prefix, e.g. "mid height"
For i = 1 To UBound(pref)
  hPos = Len(pref(i))
  Debug.Print allWords
  Debug.Print vbTab & allPreWords
  If includeNumbers = True Then
    schStr = "<" & pref(i) & " [0-9a-z]{2,}"
  Else
    schStr = "<" & pref(i) & " [a-z]{2,}"
  End If
Debug.Print pref(i)
  Set rng = ActiveDocument.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "<" & pref(i) & " [0-9a-z]{2,}"
    .Wrap = wdFindStop
    .Replacement.Text = ""
    .Forward = True
    .MatchWildcards = True
    .Execute
  End With
  
  Do While rng.Find.Found = True
    wd = rng.Text
    If InStr(wd, " ") = 0 Then wd = Left(wd, hPos) _
         & " " & Mid(wd, hPos + 1)
    wd = Replace(wd, " ", "-")
    If InStr(allPreWords, "," & wd & ",") = 0 And _
         InStr(allWords, "," & wd & ",") = 0 Then
      myPairs = myPairs + 1
      pr(myPairs) = wd
      allPreWords = allPreWords & wd & ","
      If myPairs Mod 10 = 0 Then
        If doingSeveralMacros = False Then _
             Debug.Print wd, myPairs
        StatusBar = wd & "     " & myPairs
      End If
    End If
    DoEvents
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
  Loop
Next i
halfTime = Timer

' Count the frequencies
Selection.HomeKey Unit:=wdStory
Selection.TypeText vbCr & vbCr
Selection.HomeKey Unit:=wdStory
ActiveDocument.Paragraphs(1).Style = _
     ActiveDocument.Styles(wdStyleHeading1)

allText = " " & ActiveDocument.Range.Text & " "
     
' At this point, change all "^p" to "^p "
' all punctuation to " "
chs = " , . ! : ; [ ] { } ( ) / \ + "
chs = chs & ChrW(8220) & " "
chs = chs & ChrW(8221) & " "
chs = chs & ChrW(8201) & " "
chs = chs & ChrW(8222) & " "
chs = chs & ChrW(8217) & " "
chs = chs & ChrW(8216) & " "
chs = chs & ChrW(8212) & " "
chs = chs & ChrW(8722) & " "
chs = chs & vbCr & " "
chs = chs & vbTab & " "

' To force space at start; no space at end
' i.e. one space for each character that
' needs changing to a space
chs = " " & chs & " "
chs = Replace(chs, "  ", " ")
chs = Replace(chs, "  ", " ")
chs = Left(chs, Len(chs) - 1)

chars = Split(chs, " ")
For i = 1 To UBound(chars)
  allText = Replace(allText, chars(i), " ")
Next i
allText = Replace(allText, "  ", " ")

cnt = Len(allText)
For i = 1 To myPairs
  totFinds = 0
  thisFind = ""
  Set rng = ActiveDocument.Content
  myTot = rng.End
  wdHyph = pr(i)
  wd = Replace(wdHyph, "-", "")
  For j = 1 To 4
    Select Case j
      Case 1: schWd = wdHyph
      Case 2: schWd = Replace(wdHyph, "-", " ")
      Case 3: schWd = wd
      Case 4: schWd = Replace(wdHyph, "-", ChrW(8211))
    End Select
    sc = " " & schWd & " "
    myCount = Len(Replace(allText, sc, sc & "!")) - cnt
    If myCount > 0 Then
      totFinds = totFinds + 1
      Selection.HomeKey Unit:=wdStory
      thisFind = thisFind & schWd & " . ." & _
           Str(myCount) & ":"
    Else
      thisFind = thisFind & ":"
    End If
    DoEvents
  Next j
  If (myPairs - i) Mod 10 = 0 Then
    If doingSeveralMacros = False Then _
         Debug.Print "To go:  ", myPairs - i
  End If
  If Len(thisFind) > 8 Then myResult = myResult & "%" & _
       wd & "%" & thisFind & "!"
Next i
myResult = Replace(myResult, ":!", vbCr)
myResult = Replace(myResult, ":", vbTab)
Selection.WholeStory
Selection.Delete
Set rng = ActiveDocument.Content
rng.InsertAfter myResult
Selection.Sort SortOrder:=wdSortOrderAscending
Selection.HomeKey Unit:=wdStory

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "%*%"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchCase = False
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

Selection.HomeKey Unit:=wdStory
Selection.TypeText "Hyphenation use"
startTable = Selection.End + 1
ActiveDocument.Paragraphs(1).Style = _
     ActiveDocument.Styles(wdStyleHeading1)
Selection.Start = startTable
Selection.End = ActiveDocument.Range.End
Selection.ConvertToTable Separator:=wdSeparateByTabs

Set tb = ActiveDocument.Tables(1)
For i = 1 To tb.Rows.count
  num = 0
  For j = 1 To 4
    If Len(tb.Cell(i, j).Range.Text) > 2 Then num = num + 1
  Next j
  If num = 1 Then
    For j = 1 To 4
      tb.Cell(i, j).Range.Font.ColorIndex = lighterColour
    Next j
  End If
Next i

Set tb = ActiveDocument.Tables(1)
For i = 1 To tb.Rows.count
  For j = 1 To 4
    hyphPos = 0
    txt = tb.Cell(i, j).Range.Text
    hyphPos = InStr(txt, "-")
    dashPos = InStr(txt, ChrW(8211))
    tstText = txt
    If hyphPos + dashPos > 0 Then
      tstText = "," & Left(txt, hyphPos + dashPos _
           - 1) & ","
      If InStr(myList, tstText) > 0 Then
        tb.Cell(i, j).Range.Font.ColorIndex = wdBlue
      End If
    Else
      For k = 1 To UBound(pref)
        If InStr("," & txt, "," & pref(k)) > 0 Then
          tb.Cell(i, j).Range.Font.ColorIndex = wdBlue
        End If
      Next k
    End If
  Next j
Next i

For i = 1 To tb.Rows.count
  S = 0
  If Len(tb.Cell(i, 1).Range.Text) > 2 Then S = S + 1
  If Len(tb.Cell(i, 3).Range.Text) > 2 Then S = S + 1
  If Len(tb.Cell(i, 4).Range.Text) > 2 Then S = S + 1
  If Len(tb.Cell(i, 2).Range.Text) > 2 And _
       Len(tb.Cell(i, 4).Range.Text) > 2 Then S = 2
  If S > 1 Then
    For j = 1 To 4
      tb.Cell(i, j).Range.Font.ColorIndex = wdRed
    Next j
  End If
  If InStr(tb.Cell(i, 1).Range.Text, "ly-") > 0 And _
       Len(tb.Cell(i, 2).Range.Text) > 2 Then
    For j = 1 To 4
      tb.Cell(i, j).Range.Font.ColorIndex = wdRed
    Next j
  End If
Next i

allText = ActiveDocument.Content
For Each myCell In tb.Range.Cells
  myText = myCell.Range.Text
  Set rng = myCell.Range.Duplicate
  rng.End = rng.Start + 1
  myColour = rng.Font.ColorIndex
  i = InStr(myText, " . .")
  If myColour = lighterColour And i > 2 Then
    myWord = Left(myText, i - 1)
    If Right(myWord, 1) = "s" Then
      mySingular = Left(myText, i - 2)
      If InStr(allText, mySingular & " . .") > 0 Then _
        myCell.Range.Font.Color = wdColorAutomatic
      myTest = Replace(mySingular, "-", "")
      If InStr(allText, mySingular & " . .") > 0 Then _
        myCell.Range.Font.Color = wdColorAutomatic
      myTest = Replace(mySingular, "-", " ")
      If InStr(allText, myTest & " . .") > 0 Then _
        myCell.Range.Font.Color = wdColorAutomatic
    End If
    If InStr(allText, myWord & "s . .") > 0 Then _
      myCell.Range.Font.Color = wdColorAutomatic
    myTest = Replace(myWord, "-", "")
    If InStr(allText, myTest & "s . .") > 0 Then _
      myCell.Range.Font.Color = wdColorAutomatic
    myTest = Replace(myText, "-", " ")
    If InStr(allText, myWord & "s . .") > 0 Then _
      myCell.Range.Font.Color = wdColorAutomatic
  End If
Next myCell

tb.Style = "Table Grid"
tb.AutoFitBehavior (wdAutoFitContent)
If deleteTableBorders = True Then
  tb.Borders(wdBorderTop).LineStyle = wdLineStyleNone
  tb.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
  tb.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
  tb.Borders(wdBorderRight).LineStyle = wdLineStyleNone
  tb.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
  tb.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
End If
Selection.HomeKey Unit:=wdStory

timNow = Timer
Application.ScreenUpdating = True

If doingSeveralMacros = False Then
  timGone = timNow - strttime
  Beep
  myTime = Timer
  Do
  Loop Until Timer > myTime + 0.2
  Beep
  m = Int(timGone / 60)
  S = Int(timGone) - m * 60
  timeAll = "Time:  " & Trim(Str(m)) & " m " & _
       Trim(Str(S)) & " s"
  Selection.HomeKey Unit:=wdStory
  numPairs = ActiveDocument.Tables(1).Rows.count
  MsgBox "Items:  " & Trim(Str(numPairs)) & vbCr & vbCr _
       & timeAll
Else
  FUT.Activate
End If
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub

Sub ListAlyse()
' Paul Beverley - Version 11.11.20
' Makes a list of all the 'list' items - then you can analyse them!

Set orig = ActiveDocument.Content

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

If doingSeveralMacros = False Then
  myResponse = MsgBox("    ListAlyse" & vbCr & vbCr & _
       "Analyse this document?", vbQuestion _
       + vbYesNoCancel, "ListAlyse")
  If myResponse <> vbYes Then Exit Sub
End If

Documents.Add
' Selection.Text = orig.Text
Selection.FormattedText = orig.FormattedText
ActiveDocument.ConvertNumbersToText

Set rng = ActiveDocument.Content
Set wasRng = ActiveDocument.Paragraphs(1)

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^13{2,}"
  .Wrap = wdFindContinue
  .Forward = True
  .Replacement.Text = "^p"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With

numTables = ActiveDocument.Tables.Count
If numTables > 0 Then
  For i = numTables To 1 Step -1
   ActiveDocument.Tables(i).Delete
  Next i
End If

Set rng = ActiveDocument.Content
rng.Font.Underline = True
inList = False
For Each myPara In ActiveDocument.Paragraphs
  isItem = False
  ' Bullet
  If myPara.Range.Characters(1) = ChrW(8226) Then isItem = True
  ' Weird bullets from auto-lists
  Debug.Print myPara.Range.Characters(1)
  Debug.Print Asc(myPara.Range.Characters(1))
  myPara.Range.Select
  If Asc(myPara.Range.Characters(1)) = 63 Then isItem = True
  ' en dash
  If myPara.Range.Characters(1) = ChrW(8211) Then isItem = True
  ' numbered line
  If Val(myPara.Range.Words(1)) > 0 Then isItem = True
  ' a. or a)
  If InStr("abcdefghijkl", myPara.Range.Words(1)) > 0 Then isItem = True
  If isItem Then
    myPara.Range.Font.Underline = False
    wasRng.Font.Underline = False
    inList = True
  Else
    If inList = True Then
      wasRng.InsertAfter vbCr
      wasRng.Start = wasRng.Start - 1
      wasRng.Font.Underline = False
      wasRng.Select
      inList = False
    End If
  End If
  DoEvents
  Set wasRng = myPara.Range.Duplicate
Next myPara

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Font.Underline = True
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With

Selection.HomeKey Unit:=wdStory
If doingSeveralMacros = True Then
  FUT.Activate
Else
  Beep
End If
End Sub

Sub ProperNounAlyse()
' Paul Beverley - Version 17.02.24
' Analyses similar proper nouns

minLengthCheck = 4

includeAcronyms = True

ignoreWords = "The This There Those Their They Then These That"

similarChars = "bb,b; b,p; sch,sh; ch,sh; c,k; ph,f; ss,z; s,z;" & _
               " mp,m; ll,l; nn,n; nd,n; nt,n;"

' With non-English languages, you might need to make this False
ignorePlurals = True

myScreenOff = True
Set rng = Selection.Range.Duplicate
rng.End = rng.Start + 1
myLanguage = Languages(rng.LanguageID).NameLocal
Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") + _
     InStr(FUT.Name, "Document") > 0)
If doingSeveralMacros = False Then
  myResponse = MsgBox("    ProperNounAlyse" & vbCr & vbCr & _
       "Analyse this document?", vbQuestion _
       + vbYesNoCancel, "ProperNounAlyse")
  If myResponse <> vbYes Then Exit Sub
End If
If myScreenOff = True Then
  Application.ScreenUpdating = False
  On Error GoTo ReportIt
End If

myDummy = ChrW(222)
For i = 1 To 100
  spcs = " " & spcs
Next i

dummyText = ChrW(197) & "zzzx "
For i = 65 To 90
  dummyText = dummyText & ChrW(i) & "zzzz "
Next i

checkFinalLetters = True
' checkFinalLetters = False
' Grey on word only
thisHighlight = wdGray25

doMissingLetter = True
' doMissingLetter = False
' Bold And blue

switchTest = True
' switchTest = False
' double strikethrough

doSimilarLetters = True
' doSimilarLetters = False
' various highlight colours + underline

doVowelTest = True
' doVowelTest = False
' various highlight colours + italic

' These last two tests cycle through these colours:
maxCol = 6
ReDim myCol(maxCol) As Integer
myCol(1) = wdYellow
myCol(2) = wdBrightGreen
myCol(3) = wdTurquoise
myCol(4) = wdRed
myCol(5) = wdPink
myCol(6) = wdGray25
colcode = 0

oldColour = Options.DefaultHighlightColorIndex
Options.DefaultHighlightColorIndex = wdGray25
leadDots = " . . . "
title1 = "Proper noun list"
title2 = "Proper noun queries"
CR = vbCr: CR2 = CR & CR
convCharsUC = "AAAAAAA.EEEEIIII..OOOOO.OUUUU" & _
     "...aaaaaaa.eeeeiiiio.ooooo.ouuuu......"
convCharsLC = LCase(convCharsUC)
timeStart = Timer

' collect notes text, if any
endText = ""
footText = ""
If ActiveDocument.Endnotes.Count > 0 Then
  endText = ActiveDocument.StoryRanges(wdEndnotesStory).Text
End If
If ActiveDocument.Footnotes.Count > 0 Then
  footText = ActiveDocument.StoryRanges(wdFootnotesStory).Text
End If

' collect text in all the textboxes (if any)
sh = ActiveDocument.Shapes.Count
If sh > 0 Then
  ReDim shText(sh)
  i = 0
  For Each shp In ActiveDocument.Shapes
    If shp.Type <> 24 And shp.Type <> 3 Then
      If shp.TextFrame.HasText Then
        i = i + 1
        shText(i) = shp.TextFrame.TextRange.Text
      End If
    End If
  Next
  shCount = i
End If

' Create various documents
Set rng = ActiveDocument.Content
Documents.Add
Set firstDoc = ActiveDocument
Set fnl = ActiveDocument.Content

Documents.Add
Set tempDoc = ActiveDocument
Set tmp = ActiveDocument.Content

Documents.Add
Set allText = ActiveDocument
Selection.TypeText dummyText & vbCr
Selection.FormattedText = rng.FormattedText
Selection.Collapse wdCollapseEnd

' Add notes + shape text
Selection.TypeText endText & CR & footText & CR
If shCount > 0 Then
  For i = 1 To shCount
    Selection.TypeText shText(i) & CR
  Next i
End If
Selection.HomeKey Unit:=wdStory

Set rng = allText.Content
rng.Revisions.AcceptAll
DoEvents
StatusBar = spcs & "Preparing copied file - 1"
DoEvents

' Delete struck-through text
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .MatchWildcards = False
  .Font.StrikeThrough = True
  .Replacement.Text = " "
  .Execute Replace:=wdReplaceAll
End With

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "["
  .MatchWildcards = False
  .Replacement.Text = " "
  .Execute Replace:=wdReplaceAll
End With

' Remove strange unicode characters
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[" & ChrW(&HA000) & "-" & ChrW(&HD6FF) & "]{1,}"
  .MatchWildcards = True
  .Replacement.Text = " "
  .Execute Replace:=wdReplaceAll
End With
DoEvents
StatusBar = spcs & "Preparing copied file - 2"
DoEvents

' Cut all and replace as pure text
Set rng = allText.Content
tmp.FormattedText = rng.FormattedText
rng.Text = tmp.Text
tmp.Delete
DoEvents
StatusBar = spcs & "Preparing copied file - 3"

' Use qqq for apostrophe
Set rng = allText.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "n" & ChrW(8217) & "t"
  .MatchWildcards = False
  .Replacement.Text = "nqqqt"
  .Execute Replace:=wdReplaceAll
End With

' Use qq for apostrophe
With rng.Find
  .Text = "O'"
  .MatchCase = True
  .Replacement.Text = "Oqqq"
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

' Find initial cap words
DoEvents
StatusBar = spcs & "Preparing copied file - 4"
DoEvents
myChopNum = minLengthCheck - 2
If myChop < 1 Then myChop = 1
myChop = Trim(Str(myChopNum))
myFind = "<[A-Z][a-z][a-zA-Z]{" & myChop & ",}"
If includeAcronyms = True Then myFind = _
     "<[A-Z][a-zA-Z][a-zA-Z]{" & myChop & ",}"
Set rng = allText.Range
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = myFind
  .MatchWildcards = True
  .MatchCase = True
  .Replacement.Text = "^&"
  .Replacement.Highlight = True
  .Replacement.Font.StrikeThrough = True
  .Execute Replace:=wdReplaceAll
End With

' Delete all non-strikethrough words
DoEvents
StatusBar = spcs & "Preparing copied file - 5"
DoEvents

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Font.StrikeThrough = False
  .MatchWildcards = False
  .MatchCase = True
  .Replacement.Text = "^p"
  .Execute Replace:=wdReplaceAll
End With

' Delete the unwanted "proper nouns"
DoEvents
StatusBar = spcs & "Preparing copied file - 6"
igWords = Split(Trim(ignoreWords), " ")

For Each wd In igWords
  Set rng = allText.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = wd & "^p"
    .Wrap = wdFindContinue
    .Replacement.Text = ""
    .MatchCase = True
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With
  DoEvents
Next wd

StatusBar = spcs & "Sorting whole file"
DoEvents
i = 0

For ch = 65 To 90
  For Each myPara In allText.Paragraphs
    If Asc(myPara.Range) = ch Then
      DoEvents
      myPara.Range.Font.StrikeThrough = False
      tmp.InsertAfter myPara.Range.Text
    End If
  Next myPara
  tmp.InsertAfter Text:="Zzzzz" & CR

  Set rng = tempDoc.Content
  rng.Sort SortOrder:=wdSortOrderAscending, CaseSensitive:=True

  ' delete initial blank line
  If Len(tempDoc.Paragraphs(1)) < 3 Then _
       tempDoc.Paragraphs(1).Range.Delete

  ' Create a frequency for each highlighted word
  thisWord = ""
  myCount = 0
  For Each myPara In tempDoc.Paragraphs
    Set rng = myPara.Range.Words(1)
    DoEvents
    nextWord = rng
    If nextWord <> thisWord Then
    ' This is a new word
      If Len(thisWord) > 1 Then
        fnl.InsertAfter Text:=thisWord _
             & leadDots & Trim(Str(myCount)) & CR
      End If
      thisWord = nextWord
      myCount = 1
    Else
      myCount = myCount + 1
    End If
    If nextWord = "Zzzzz" Then Exit For
    i = i + 1:
    If i Mod 400 = 4 Then
      DoEvents
      prmt = Left(thisWord, 1) & " "
      prmt = prmt & prmt & prmt & prmt
      StatusBar = spcs & _
           "Preparing words for frequency list - " & prmt
      DoEvents
    End If
  Next myPara

  ' Remove all words except frequency counts
  Set rng = tempDoc.Content
  rng.Delete
Next ch

' Find any unaccounted-for words, e.g. Ångstrom
For Each myPara In allText.Paragraphs
  If myPara.Range.Words(1).Font.StrikeThrough = True Then
    tmp.InsertAfter myPara.Range.Text
  End If
Next myPara

tempDoc.Close SaveChanges:=False
allText.Close SaveChanges:=False
firstDoc.Activate

' Remove blank lines
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[^13]{2,}"
  .Wrap = wdFindContinue
  .Replacement.Text = "^p"
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

' Resort case insensitively
Set rng = ActiveDocument.Content
rng.Sort SortOrder:=wdSortOrderAscending, _
     CaseSensitive:=False

' Delete rubbish from top and bottom of list
Do
  Set rng = ActiveDocument.Paragraphs(1).Range
  myLen = Len(rng.Text)
  If myLen < 10 Then
    rng.Select
    Selection.Delete
  End If
Loop Until myLen > 9
Do
  lastLine = ActiveDocument.Paragraphs.Count
  Set rng = ActiveDocument.Paragraphs(lastLine).Range
  myLen = Len(rng.Text)
  If myLen < 10 Then
    rng.Select
    Selection.Delete
  End If
Loop Until Len(rng.Text) >= 2

' Word list now has freq. count.
Do
  lastLine = ActiveDocument.Paragraphs.Count
  Set rng = ActiveDocument.Paragraphs(lastLine).Range
  myLen = Len(rng.Text)
  If myLen < 10 Then
    rng.Select
    Selection.Delete
  End If
Loop Until Len(rng.Text) >= 2

' Create another copy for doing extra tests
Set rng = ActiveDocument.Content
Documents.Add
Set finalList = ActiveDocument
finalList.Range.Text = rng.Text
Selection.HomeKey Unit:=wdStory

' Prepare data for other tests
numWords = ActiveDocument.Paragraphs.Count
For i = 1 To numWords
  aWord = ActiveDocument.Paragraphs(i).Range.Words(1)
  n = AscW(aWord)
  thisChar = ChrW(n)
  If n > 129 Then
    If n >= 217 Then aWord = Replace(aWord, thisChar, "U")
    If n >= 210 Then aWord = Replace(aWord, thisChar, "O")
    If n >= 204 Then aWord = Replace(aWord, thisChar, "I")
    If n >= 200 Then aWord = Replace(aWord, thisChar, "E")
    If n >= 192 Then aWord = Replace(aWord, thisChar, "A")
  End If
  allWords = allWords & aWord
  jmp = 100
  If i Mod jmp = 1 Then
    pq = pq + 1
    DoEvents
    StatusBar = spcs & _
         "Preparing data for other tests - 1 - " & pq
    DoEvents
  End If
Next i

' ...for the vowel test below
DoEvents
StatusBar = spcs & "Preparing data for other tests - 2"
DoEvents
noVowelWords = " " & allWords
noVowelWords = Replace(noVowelWords, " A", "_1")
noVowelWords = Replace(noVowelWords, " E", "_2")
noVowelWords = Replace(noVowelWords, " I", "_3")
noVowelWords = Replace(noVowelWords, " O", "_4")
noVowelWords = Replace(noVowelWords, " U", "_5")
noVowelWords = Replace(noVowelWords, " Y", "_6")
For k = 2 To Len(noVowelWords) - 1
  thisChar = Mid(noVowelWords, k, 1)
  n = AscW(thisChar)
  If n > 191 And n < 221 Then
    myNewChar = Mid(convCharsLC, n - 191, 1)
    If myNewChar <> "." Then noVowelWords = _
         Replace(noVowelWords, thisChar, myNewChar)
  End If
Next k
noVowelWords = Replace(noVowelWords, "a", "")
noVowelWords = Replace(noVowelWords, "e", "")
noVowelWords = Replace(noVowelWords, "i", "")
noVowelWords = Replace(noVowelWords, "o", "")
noVowelWords = Replace(noVowelWords, "u", "")
noVowelWords = Replace(noVowelWords, "y", "")
noVowelWords = Replace(noVowelWords, "A", "")
noVowelWords = Replace(noVowelWords, "E", "")
noVowelWords = Replace(noVowelWords, "I", "")
noVowelWords = Replace(noVowelWords, "O", "")
noVowelWords = Replace(noVowelWords, "U", "")
noVowelWords = Replace(noVowelWords, "Y", "")
noVowelWords = Replace(noVowelWords, "_1", " A")
noVowelWords = Replace(noVowelWords, "_2", " E")
noVowelWords = Replace(noVowelWords, "_3", " I")
noVowelWords = Replace(noVowelWords, "_4", " O")
noVowelWords = Replace(noVowelWords, "_5", " U")
noVowelWords = Replace(noVowelWords, "_6", " Y")

' ...for the similar words test
DoEvents
StatusBar = spcs & "Preparing data for other tests - 3"
DoEvents
similarAllWords = " " & LCase(allWords)
similarChars = Replace(similarChars, " ", "")
sChars = Replace(similarChars, " ", "")

Do
  commaPos = InStr(sChars, ",")
  charWas = Left(sChars, commaPos - 1)
  sChars = Mid(sChars, commaPos + 1)
  semicolonPos = InStr(sChars, ";")
  charNew = Left(sChars, semicolonPos - 1)
  sChars = Mid(sChars, semicolonPos + 1)
  similarAllWords = Replace(similarAllWords, charWas, charNew)
Loop Until Len(sChars) < 2

' Changes all the accented characters to non-accented
DoEvents
StatusBar = spcs & "Preparing data for other tests - 4"
DoEvents
sWd = similarAllWords
For k = 1 To Len(sWd) - 1
  thisChar = Mid(sWd, k, 1)
  n = AscW(thisChar)
  myNewChar = "."
  If n > 191 And n < 256 Then
    myNewChar = Mid(convCharsLC, n - 191, 1)
    If myNewChar <> "." Then sWd = Replace(sWd, _
         thisChar, myNewChar)
  End If
Next k
similarAllWords = sWd

' Catch words with only the final two letters the same
i = 0
If checkFinalLetters = True Then
  For Each myPara In ActiveDocument.Paragraphs
    gotOne = False
    myWord = Trim(myPara.Range.Words(1))
    myLen = Len(myWord)
    If myLen > 6 Then
      myTarget = "^p" & Left(myWord, myLen - 2) & "^$^$ "
      myCut = 2
    Else
      myTarget = "^p" & Left(myWord, myLen - 1) & "^$ "
      myCut = 1
    End If
    Set rng = ActiveDocument.Content
    rng.Start = myPara.Range.End - 3
    rng.Collapse wdCollapseStart
    With rng.Find
      .Replacement.ClearFormatting
      .ClearFormatting
      .Text = myTarget
      .Replacement.Text = ""
      .Forward = True
      .MatchCase = True
      .MatchWildcards = False
      .Wrap = wdFindStop
    End With
    rng.Find.Execute
    Do While rng.Find.Found
      gotOne = True
      rng.MoveStart 1
      rng.End = rng.Start + myLen - myCut
      rng.HighlightColorIndex = thisHighlight
      rng.Font.Bold = True
      rng.Find.Execute
    Loop
    If gotOne = True Then
      Set rng = myPara.Range.Words(1)
      rng.End = rng.Start + myLen - myCut
      rng.HighlightColorIndex = thisHighlight
      rng.Font.Bold = True
    End If
    i = i + 1
    If i Mod 100 = 1 Then
      DoEvents
      StatusBar = spcs & "Doing test (5) on " & myWord
      DoEvents
    End If
  Next myPara
End If

If doMissingLetter = True Then
' Start of test
  doneWords = ""
  doneSimilarWords = ""
  McList = ""

  For i = 1 To ActiveDocument.Paragraphs.Count - 1
    myWord = ActiveDocument.Paragraphs(i).Range.Words(1)
    n = AscW(myWord)
    thisChar = ChrW(n)
    myNewChar = "."
  ' Changes the capital letter, if a vowel
    If n > 191 And n < 221 Then
      myNewChar = Mid(convCharsUC, n - 191, 1)
      If myNewChar <> "." Then myWord = Replace(myWord, _
           thisChar, myNewChar)
    End If

    If i Mod 50 = 1 Then
      DoEvents
      StatusBar = spcs & "Other tests (4) on " & myWord
      DoEvents
    End If
    testWords = Replace(allWords, myWord, "")
    captestLetters = Left(myWord, 1)

  ' Check if word reappears with one letter missing (1)
    For k = 2 To Len(myWord) - 1
      testWord = " " & Left(myWord, k - 1) & Mid(myWord, k + 1)
      wordPos = InStr(allWords, testWord)
      If wordPos > 0 Then
        lastLetter = Mid(myWord, Len(myWord) - 1, 1)
      ' but not "s" at the end, unless it's a spelling error
        If lastLetter = "s" Then
          ignoreIt = (Application.CheckSpelling(myWord, _
          MainDictionary:=myLanguage) = True)
        Else
          ignoreIt = False
        End If
        If ignoreIt = False And ignorePlurals = True Then
          colcode = (colcode + 1) Mod maxCol
          thisCol = myCol(colcode + 1)

          ' mark the pair
          leftBit = Left(allWords, InStr(allWords, testWord) _
               + Len(testWord) - 1)
          j = Len(leftBit) - Len(Replace(leftBit, " ", ""))
          Set rng = ActiveDocument.Paragraphs(i).Range
          rng.HighlightColorIndex = thisCol
          rng.Font.Bold = True
          rng.Font.Color = wdColorBlue
          Set rng = ActiveDocument.Paragraphs(j).Range
          rng.HighlightColorIndex = thisCol
          rng.Font.Bold = True
          rng.Font.Color = wdColorBlue
        End If
      End If
    Next k

    If Left(myWord, 2) = "Mc" Or Left(myWord, 3) = "Mac" Or _
         Left(myWord, 3) = "Mag" Then
      McList = McList & ActiveDocument.Paragraphs(i).Range
    End If
  Next i
End If

If doSimilarLetters = True Then
  doneWords = ""
  doneSimilarWords = ""

  For i = 1 To ActiveDocument.Paragraphs.Count - 1
    myWord = ActiveDocument.Paragraphs(i).Range.Words(1)
    n = AscW(myWord)
    thisChar = ChrW(n)
    myNewChar = "."
   ' Changes the capital letter, if a vowel
    If n > 191 And n < 221 Then
      myNewChar = Mid(convCharsUC, n - 191, 1)
      If myNewChar <> "." Then myWord = Replace(myWord, _
           thisChar, myNewChar)
    End If
    If i Mod 50 = 1 Then
      DoEvents
      StatusBar = spcs & "Other tests (3) on " & myWord
      DoEvents
    End If
    testWords = Replace(allWords, myWord, "")
    captestLetters = Left(myWord, 1)

' check similar spellings: Perutz/Peruts or Chebyshev/Chevychev
    similarWord = " " & LCase(myWord)
    sChars = similarChars
    Do
      commaPos = InStr(sChars, ",")
      charWas = Left(sChars, commaPos - 1)
      sChars = Mid(sChars, commaPos + 1)
      semicolonPos = InStr(sChars, ";")
      charNew = Left(sChars, semicolonPos - 1)
      sChars = Mid(sChars, semicolonPos + 1)
      similarWord = Replace(similarWord, charWas, charNew)
    Loop Until Len(sChars) < 2
    ' Changes all the accented characters to non-accented
    For k = 1 To Len(myWord) - 1
      thisChar = Mid(myWord, k, 1)
      n = AscW(thisChar)
      If n > 191 And n < 256 Then
        myNewChar = Mid(convCharsUC, n - 191, 1)
        If myNewChar <> "." Then myWord = Replace(myWord, _
             thisChar, myNewChar)
      End If
    Next k
    similarAllWords = Mid(similarAllWords, Len(similarWord))
    theseWords = similarAllWords
    If InStr(doneSimilarWords, similarWord) = 0 And _
          InStr(theseWords, similarWord) > 0 Then
      colcode = (colcode + 1) Mod maxCol
      thisCol = myCol(colcode + 1)
      Set rng = ActiveDocument.Paragraphs(i).Range
      rng.HighlightColorIndex = thisCol
      rng.Font.Underline = True
      doneSimilarWords = doneSimilarWords & similarWord
      ' search through all the following words
      theseWords = similarAllWords
      For j = 1 To numWords - i
        spPos = InStr(Trim(theseWords) & " ", " ")
        If Left(theseWords, spPos + 1) = similarWord Then
          Set rng = ActiveDocument.Paragraphs(i + j).Range
          rng.HighlightColorIndex = thisCol
          rng.Font.Underline = True
        End If
        theseWords = Mid(theseWords, spPos + 1)
        capThisLetter = Mid(theseWords, 2, 1)
        If capThisLetter <> LCase(captestLetters) Then Exit For
      Next j
    End If
  Next i
End If

If switchTest = True Then
  doneWords = ""
  doneSimilarWords = ""
  McList = ""
  For i = 1 To ActiveDocument.Paragraphs.Count - 1
    myWord = ActiveDocument.Paragraphs(i).Range.Words(1)
    n = AscW(myWord)
    thisChar = ChrW(n)
    myNewChar = "."
   ' Changes the capital letter, if a vowel
    If n > 191 And n < 221 Then
      myNewChar = Mid(convCharsUC, n - 191, 1)
      If myNewChar <> "." Then myWord = Replace(myWord, _
           thisChar, myNewChar)
    End If
    If i Mod 50 = 1 Then
      DoEvents
      StatusBar = spcs & "Other tests (2) on " & myWord
      DoEvents
    End If
    testWords = Replace(allWords, myWord, "")
    captestLetters = Left(myWord, 1)

' check for switched chars
    wordLen = Len(myWord) - 1
    For k = 1 To Len(myWord) - 3
      otherWord = Left(myWord, k) & Mid(myWord, k + 2, 1) & _
            Mid(myWord, k + 1, 1) & Mid(myWord, k + 3)
      wordPos = InStr(testWords, otherWord)
      If wordPos > 0 Then
      ' Find the position of the matching word
        matchWord = Mid(testWords, wordPos, Len(myWord))
        leftBit = Left(allWords, InStr(allWords, matchWord) + 1)
        j = Len(leftBit) - Len(Replace(leftBit, " ", "")) + 1
        ActiveDocument.Paragraphs(i).Range.Font.DoubleStrikeThrough _
             = True
        ActiveDocument.Paragraphs(i).Range.HighlightColorIndex _
             = thisCol
        ActiveDocument.Paragraphs(j).Range.Font.DoubleStrikeThrough _
             = True
        ActiveDocument.Paragraphs(j).Range.HighlightColorIndex _
             = thisCol
      End If
    Next k
  Next i
End If

If doVowelTest = True Then
  doneWords = ""
  doneSimilarWords = ""
  McList = ""
  For i = 1 To ActiveDocument.Paragraphs.Count - 1
    myWord = ActiveDocument.Paragraphs(i).Range.Words(1)
    n = AscW(myWord)
    thisChar = ChrW(n)
    myNewChar = "."
   ' Changes the capital letter, if a vowel
    If n > 191 And n < 221 Then
      myNewChar = Mid(convCharsUC, n - 191, 1)
      If myNewChar <> "." Then myWord = Replace(myWord, _
           thisChar, myNewChar)
    End If
    If i Mod 50 = 1 Then
      DoEvents
      StatusBar = spcs & "Other tests (1) on " & myWord
      DoEvents
    End If
    testWords = Replace(allWords, myWord, "")
    captestLetters = Left(myWord, 1)

    ' check if there's a word with different vowels
    otherWord = " " & Replace(myWord, "a", "")
    otherWord = Replace(otherWord, "e", "")
    otherWord = Replace(otherWord, "i", "")
    otherWord = Replace(otherWord, "o", "")
    otherWord = Replace(otherWord, "u", "")
    otherWord = Replace(otherWord, "y", "")

    ' Delete all the accented characters
    For k = 3 To Len(otherWord) - 1
      thisChar = Mid(otherWord, k, 1)
      n = AscW(thisChar)
      If InStr("AEIOUY", thisChar) > 0 Then
        otherWord = Left(otherWord, k - 1) & "=" & Mid(otherWord, k + 1)
      Else
        If n > 191 And n < 221 Then
          myNewChar = Mid(convCharsUC, n - 191, 1)
          If myNewChar <> "." Then
            otherWord = Replace(otherWord, thisChar, "=")
          End If
        End If
      End If
    Next k
    otherWord = Replace(otherWord, "=", "")

' otherWord is now the word under test (vowel-less)
    otherWord = Replace(otherWord, ".", "")
    noVowelWords = Mid(noVowelWords, Len(otherWord))
    If Left(noVowelWords, 1) <> " " Then noVowelWords = _
         " " & noVowelWords
    theseWords = noVowelWords
    
    wordPos = InStr(noVowelWords, otherWord)
    If InStr(doneWords, otherWord) = 0 And wordPos > 0 Then
      colcode = (colcode + 1) Mod maxCol
      thisCol = myCol(colcode + 1)
      Set rng = ActiveDocument.Paragraphs(i).Range
      rng.HighlightColorIndex = thisCol
      rng.Font.Italic = True
      doneWords = doneWords & otherWord
      For j = 1 To numWords - i
        spPos = InStr(Trim(theseWords) & " ", " ")
        firstWord = Left(theseWords, spPos + 1)
        theseWords = Mid(theseWords, spPos + 1)
        If firstWord = otherWord Then
          Set rng = ActiveDocument.Paragraphs(i + j).Range
          rng.HighlightColorIndex = thisCol
          rng.Font.Italic = True
        End If
        capThisLetter = Mid(theseWords, 2, 1)
        If capThisLetter > "" And capThisLetter <> _
             captestLetters Then Exit For
      Next j
    End If
  Next i
End If

finishOff:
Selection.EndKey Unit:=wdStory
Selection.TypeText CR2 & McList

Selection.HomeKey Unit:=wdStory
Selection.TypeText title1 & CR
Do
  Selection.Expand wdParagraph
  If Len(Selection) < 3 Or LCase(Selection) = _
       UCase(Selection) Then Selection.Delete
Loop Until LCase(Selection) <> UCase(Selection)
Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
Selection.Style = ActiveDocument.Styles(wdStyleHeading1)

' Restore apostrophes
Set rng = finalList.Range
With rng.Find
  .Text = "qqq"
  .MatchCase = False
  .Replacement.Text = "'"
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

' Find first highlight
Set rng = finalList.Content
With rng.Find
  .Text = "Zzzzz"
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = False
  .Execute Replace:=wdReplaceOne
End With
Set rng = finalList.Content
With rng.Find
  .Text = ""
  .Highlight = True
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = False
  .Execute
End With

rng.Select
Selection.Collapse wdCollapseStart
Set finalList = ActiveDocument

firstDoc.Activate

' Find sets of sounds-like words
StatusBar = spcs & "Sounds-like tests"
k = 0
For Each myPara In ActiveDocument.Paragraphs
  myWord = Trim(myPara.Range.Words(1))
  k = k + 1
  If k Mod 40 = 1 Then
    DoEvents
    StatusBar = spcs & "Sounds-like test: " & myWord
    DoEvents
  End If
  hasAccent = False
  For i = 1 To Len(myWord)
    ascChar = AscW(Mid(myWord, i))
    If ascChar > 128 Or ascChar = Asc("?") Then hasAccent = True
  Next i

' Go and find the first sounds-like word
  initLetter = Left(myWord, 1)
  If Len(myWord) > 2 And myPara.Range.HighlightColorIndex > 0 And _
       hasAccent = False And InStr(allSets, myWord & leadDots) _
       = 0 Then
    Set rng = ActiveDocument.Content
    Do
      With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = myWord
        .Wrap = wdFindStop
        .Replacement.Text = ""
        .MatchWildcards = False
        .MatchSoundsLike = True
        .Execute
      End With
      Set myPara = rng.Paragraphs(1).Range
      rng.Collapse wdCollapseEnd
    Loop Until Left(myPara, 1) = initLetter
    setOfWords = myPara
    gottaSet = False
    rng.Collapse wdCollapseEnd
    rng.Find.Execute
    Do While rng.Find.Found = True
      Set myPara = rng.Paragraphs(1).Range
      If Left(myPara, 1) = initLetter Then
        gottaSet = True
        setOfWords = setOfWords & myPara
      End If
      rng.Collapse wdCollapseEnd
      rng.Find.Execute
    Loop
    If gottaSet = True Then allSets = allSets & setOfWords & CR
  End If
Next myPara

Selection.WholeStory
If Len(allSets) < 2 Then
  Selection.TypeText "None found with this test"
Else
  Selection.TypeText allSets
End If
Selection.HomeKey Unit:=wdStory
Selection.TypeText "Proper nouns by sound" & CR
Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
Selection.HomeKey Unit:=wdStory

Set rng = ActiveDocument.Content
rng.HighlightColorIndex = 0
rng.Copy

ActiveDocument.Close SaveChanges:=False
finalList.Activate
' Remove highlighting from second half of words
' that are only case changes of one another
totParas = ActiveDocument.Paragraphs.Count
For i = 1 To totParas - 1
  a = Trim(ActiveDocument.Paragraphs(i).Range.Words(1))
  b = Trim(ActiveDocument.Paragraphs(i + 1).Range.Words(1))
  a = Mid(a, 2)
  b = Mid(b, 2)
  If LCase(a) = LCase(b) And Len(a) > 2 Then
    If (UCase(a) = a And LCase(b) = b) Or (UCase(b) = b And _
         LCase(a) = a) Then
      ActiveDocument.Paragraphs(i).Range.Words(1).HighlightColorIndex = 0
      ActiveDocument.Paragraphs(i + 1).Range.Words(1).HighlightColorIndex _
           = 0
    End If
  End If
  If i Mod 50 = 0 Then
    DoEvents
    StatusBar = spcs & "Final checks: " & totParas - i
    DoEvents
  End If
Next i

myOnames = ""
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^13O[!a-z]"
  .Wrap = wdFindStop
  .Replacement.Text = ""
  .Forward = True
  .MatchSoundsLike = False
  .MatchWildcards = True
  .Execute
End With

Do While rng.Find.Found = True
  rng.Collapse wdCollapseEnd
  rng.Expand wdWord
  wd = Mid(rng.Text, 3)
  rng.Expand wdParagraph
  pa = rng.Text
  Set rng2 = ActiveDocument.Content
  With rng2.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "^13" & wd
    .Wrap = wdFindStop
    .Replacement.Text = ""
    .Forward = True
    .MatchWildcards = True
    .Execute
  End With
  If rng2.Find.Found Then
    rng2.Collapse wdCollapseEnd
    rng2.Expand wdParagraph
    pa2 = rng2.Text
    myOnames = myOnames & pa2 & pa & vbCr
  End If
  rng.Collapse wdCollapseEnd
  rng.End = rng.End - 2
  rng.Find.Execute
Loop
If myOnames > "" Then
  Selection.EndKey Unit:=wdStory
  Selection.TypeText "Possible O'<something> errors" & vbCr
  Selection.MoveUp , 1
  Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
  Selection.EndKey Unit:=wdStory
  Selection.TypeText myOnames
  Selection.HomeKey Unit:=wdStory
End If

Set rng = ActiveDocument.Content
finalList.Activate
Selection.EndKey Unit:=wdStory
Selection.TypeText vbCr & vbCr & vbCr
Selection.Paste
Selection.HomeKey Unit:=wdStory

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = myDummy
  .Wrap = wdFindContinue
  .Replacement.Text = " "
  .Forward = True
  .MatchCase = False
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^$zzz^$" & leadDots & "1" & vbCr
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchCase = False
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
End With

' Clear clipboard
Set rng = ActiveDocument.Content
rng.End = 2
rng.Copy
Set finalList = ActiveDocument

StatusBar = "Creating queries list"
Set rng = ActiveDocument.Content
Documents.Add
Selection.FormattedText = rng.FormattedText
Set queriesDoc = ActiveDocument

ActiveDocument.Paragraphs(1).Range.Delete
Set rng = ActiveDocument.Content
rng.Font.StrikeThrough = True
For Each myPara In ActiveDocument.Paragraphs
  Set ch = myPara.Range.Characters(1)
  chCol = ch.HighlightColorIndex
  If chCol > 0 Then
    myPara.Range.Font.StrikeThrough = False
  End If
  myLen = Len(myPara.Range.Text)
  If myLen > 4 Then
    If chCol > 0 Then
      myPara.Range.Font.StrikeThrough = False
    End If
    Set che = myPara.Range.Characters(myLen - 2)
    If che.HighlightColorIndex > 0 Then
      myPara.Range.Font.StrikeThrough = False
    End If
  End If
Next myPara
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Font.StrikeThrough = True
  .Wrap = wdFindContinue
  .Replacement.Text = "^p"
  .Forward = True
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Set rng = ActiveDocument.Content
rng.Font.StrikeThrough = False
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[^13]{3,}"
  .Wrap = wdFindContinue
  .Replacement.Text = "^p^p"
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With

For Each myPara In ActiveDocument.Paragraphs
  myText = myPara.Range.Text
  If Len(myText) > 4 Then
    Set ch = myPara.Range.Characters(1)
    numChars = myPara.Range.Characters.Count
    Set myEnd = myPara.Range.Characters(numChars)
    colNum = ch.HighlightColorIndex Mod 8

    If ch.Font.Bold = True Then
      myTxt = "qcqc  " & Str(colNum + 1) & "  =  zczc"
    Else
      myTxt = "qcqc zczc"
    End If

    If ch.Font.Underline > 0 And colNum > 0 Then
      myBit = "* "
      myTxt = Replace(myTxt, " =  ", "")
    Else
      myBit = ""
    End If
    myPara.Range.InsertBefore myBit & myTxt

    If ch.Font.Italic = True Then
      myEnd.InsertBefore "qpqp= " & Chr(65 + colNum)
    End If
  End If
  i = i + 1
  If i Mod 20 = 0 And Len(myText) > 4 Then
  myText = Replace(myText, vbCr, "")
  StatusBar = spcs & "Creating queries list:  " & myText
  End If
  DoEvents
Next myPara

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "\* qcqc(*)zczc"
  .Wrap = wdFindContinue
  .Replacement.Text = "* \1^t"
  .Replacement.Highlight = False
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "qcqc(*)zczc"
  .Wrap = wdFindContinue
  .Replacement.Text = "\1^t"
  .Replacement.Highlight = False
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "qpqp(*)^13"
  .Replacement.Text = "^t\1^p"
  .Replacement.Highlight = False
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "= ^$"
  .Replacement.Text = ""
  .Replacement.Font.Bold = False
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Set rng = ActiveDocument.Content
rng.Font.Bold = False
rng.Font.Italic = False
rng.Font.DoubleStrikeThrough = False
rng.Font.Underline = False
rng.Font.Color = wdColorBlack
Selection.HomeKey Unit:=wdStory
Selection.TypeText title2 & CR
Set rng = ActiveDocument.Content.Paragraphs(2).Range
If rng.Text = vbCr Then rng.Delete
Set rng = ActiveDocument.Content.Paragraphs(1).Range
rng.Style = ActiveDocument.Styles(wdStyleHeading1)

StatusBar = " "
Options.DefaultHighlightColorIndex = oldColour

lighterColour = wdGray25
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "= ^$"
  .Replacement.Text = ""
  .Replacement.Font.ColorIndex = lighterColour
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
  .Text = "^#  ="
  .Replacement.Text = ""
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Application.ScreenUpdating = True
If doingSeveralMacros = False Then
  myTime = (Int(10 * (Timer - timeStart) / 60) / 10)
  Beep
  If myTime > 0 Then MsgBox myTime & "  minutes"
End If
If InStr(FUT.Name, "zzTestFile") > 0 Then FUT.Activate
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub




Sub SpecialSortsLister()
' Paul Beverley - Version 25.11.20
' Collect all special sorts in a file

listAccentedChars = True

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)


Set rng = ActiveDocument.Range
rng.Copy
Documents.Add
Selection.Paste
Selection.WholeStory
Selection.Range.Style = ActiveDocument.Styles(wdStyleNormal)
Selection.Font.Reset
numberCmnts = ActiveDocument.Comments.Count
If numberCmnts > 0 Then ActiveDocument.DeleteAllComments
Set rng = ActiveDocument.Content

CR = vbCr
If listAccentedChars = True Then
  mainChars = "[abcdefghijklmnopqrstuvwxyz" & _
       "ABCDEFGHIJKLMNOPQRSTUVWXYZ^+ ]{1,}"
Else
  mainChars = "[a-zA-Z^+ ]{1,}"
End If

With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = mainChars
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .MatchWholeWord = False
  .MatchSoundsLike = False
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Beep
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[^m0-9^13^t,.:;\!\?^=^+\-\(\)£]{1,}"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
Beep
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "[" & ChrW(8216) & ChrW(8217) & ChrW(8221) & ChrW(8220) _
       & ChrW(8230) & ChrW(174) & ChrW(176) & "]{1,}"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
  DoEvents
End With
Beep

allSorts = ""
For Each myChar In ActiveDocument.Characters
  uCode = 0
  If myChar > "" Then uCode = AscW(myChar)
  If InStr(allSorts, myChar) = 0 And uCode > 128 Then
    allSorts = allSorts & myChar & vbTab
    Select Case uCode
      Case 160: extraBit = "non-breaking space"
      Case 176: extraBit = "degree symbol"
      Case 178: extraBit = "dodgy squared symbol"
      Case 179: extraBit = "dodgy cubed symbol"
      Case 184: extraBit = "cedilla"
      Case 186: extraBit = "masculine ordinal"
      Case 215: extraBit = "proper multiply symbol"
      Case 8194: extraBit = "en space"
      Case 8195: extraBit = "em space"
      Case 8201: extraBit = "thin space"
      Case 8222: extraBit = "German open curly quote"
      Case 8226: extraBit = "ordinary bullet"
      Case 8242: extraBit = "unicode: single prime"
      Case 8243: extraBit = "unicode: double prime"
      Case 8249: extraBit = "French open quote"
      Case 8250: extraBit = "French close quote"
      Case 8722: extraBit = "minus sign"

      Case Else: extraBit = "zczc"
    End Select
    allSorts = allSorts & extraBit & CR
    DoEvents
  End If
Next myChar
Selection.WholeStory
If allSorts = "" Then allSorts = CR & "No special sorts used"
Selection.TypeText allSorts
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^tzczc"
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .Forward = True
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With
Selection.EndKey Unit:=wdStory
Selection.TypeText CR
Selection.HomeKey Unit:=wdStory
Selection.TypeText "Special sorts used" & CR
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)
If doingSeveralMacros = False Then
  Beep
Else
  FUT.Activate
End If
End Sub

Sub SpellAlyse()
' Paul Beverley - Version 25.05.23
' Complete spellchecking system

ignoreNumbers = True

spellingListName = "SpellAlyse"
spellingFreqName = "SpellAlyse frequencies"
ignoreDoc = "zzSwitchList"

myFind = "´a,´e,¨a,¨e,¨o,¨u,ˆo"
myReplace = "á,é,ä,ë,ö,ü,ô"

CR = vbCr
CR2 = CR & CR
sp1 = " "
sp2 = sp1 & sp1

myScreenOff = True

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

Dim lossLen As Long
Dim newLen As Long

numParas = FUT.Paragraphs.count
numWords = FUT.Content.ComputeStatistics(wdStatisticWords)
myProfile = numWords / numParas

If myProfile < 1.1 Then GoTo autoCorrect

' List all possible spelling errors
Set rng = ActiveDocument.Content
rng.End = 2
thisLanguage = rng.LanguageID
langText = Languages(rng.LanguageID).NameLocal

langPrompt = "Spellcheck with " & langText & " dictionary. OK?"
If doingSeveralMacros = False Then
  myResponse = MsgBox(langPrompt, vbQuestion + vbYesNoCancel, _
       "SpellAlyse")
  If myResponse <> vbYes Then Exit Sub
End If

If myScreenOff = True Then
  Application.ScreenUpdating = False
  On Error GoTo ReportIt
End If
timeStart = Timer

allExcepts = CR
'Collect all words from all word lists
For Each myDoc In Documents
Debug.Print myDoc.Name
  pNum = myDoc.Paragraphs.count
  myNum = 3
  If pNum < 3 Then myNum = pNum
  Set rng = myDoc.Paragraphs(myNum).Range
  rng.Start = 0
  If InStr(LCase(rng.Text), "elist") > 0 Then
    For Each myPar In myDoc.Paragraphs
      myWord = Trim(myPar.Range.Words(1).Text)
      If Len(myWord) > 2 Then
        allExcepts = allExcepts & myWord & CR
        Debug.Print allExcepts
      End If
      DoEvents
    Next myPar
  End If
  DoEvents
Next myDoc
Debug.Print allExcepts

' Create a text-only copy in another file
Set rngOld = FUT.Content
Documents.Add
Set erList = ActiveDocument
Set rng = ActiveDocument.Content
rng.LanguageID = thisLanguage
rng.FormattedText = rngOld.FormattedText

numNotes = FUT.Endnotes.count
If numNotes > 0 Then
  rng.Collapse wdCollapseEnd
  rng.FormattedText = FUT.StoryRanges(wdEndnotesStory).FormattedText
End If

numNotes = FUT.Footnotes.count
If numNotes > 0 Then
  rng.Collapse wdCollapseEnd
  rng.FormattedText = FUT.StoryRanges(wdFootnotesStory).FormattedText
End If

' copy all the textboxes to the end of the text
shCount = FUT.Shapes.count
If shCount > 0 Then
  Selection.EndKey Unit:=wdStory
  For j = 1 To shCount
    Set shp = FUT.Shapes(j)
    If shp.Type <> 24 And shp.Type <> 3 Then
      If shp.TextFrame.HasText Then
        Set rng = shp.TextFrame.TextRange
        If Len(rng.Text) > 1 Then
          Selection.Text = rng.Text
          Selection.EndKey Unit:=wdStory
        End If
      End If
    End If
    DoEvents
  Next
End If

' Add a newline for safety
Selection.TypeText CR

Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Font.StrikeThrough = True
  .Wrap = wdFindContinue
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = "^2"
  .Replacement.Text = ""
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  
  .Text = "([!a-zA-Z])['" & ChrW(8216) & "]"
  .Replacement.Text = "\1"
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
End With

' Change ligature characters into character pairs
myFind = myFind & "," & ChrW(-1280) & "," & ChrW(-1279) & _
     "," & ChrW(-1278) & "," & ChrW(-1277) & "," & ChrW(-1276) _
     & "," & ChrW(185) & "," & ChrW(178) & "," & ChrW(179)
myReplace = myReplace & ",ff,fi,fl,ffi,ffl, , , , , "
fnd = Split(myFind, ",")
rpl = Split(myReplace, ",")

For i = 0 To UBound(fnd)
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = fnd(i)
    .Wrap = wdFindContinue
    .Replacement.Text = rpl(i)
    .MatchCase = False
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With
  DoEvents
Next i

Selection.HomeKey Unit:=wdStory
Selection.TypeText Text:="Preparing file for spellchecking" & CR2
ActiveDocument.Paragraphs(1).Style = _
     ActiveDocument.Styles(wdStyleHeading1)
DoEvents
Set rng = ActiveDocument.Content
With rng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
   If ignoreNumbers = True Then
     .Text = "[!a-zA-Z^13'" & ChrW(8217) _
     & ChrW(248) & "-" & ChrW(591) _
     & ChrW(697) & "-" & ChrW(703) _
     & ChrW(&H591) & "-" & ChrW(&H5FF) _
     & ChrW(7680) & "-" & ChrW(7935) & "]"
   Else
     .Text = "[!a-zA-Z0-9^13'" & ChrW(8217) _
     & ChrW(248) & "-" & ChrW(591) _
     & ChrW(697) & "-" & ChrW(703) _
     & ChrW(&H591) & "-" & ChrW(&H5FF) _
     & ChrW(7680) & "-" & ChrW(7935) & "]"
   End If
  .Replacement.Text = sp1
  .Wrap = wdFindContinue
  .MatchWildcards = True
  .Execute Replace:=wdReplaceAll
       
  .Text = "([a-zA-Z])['" & ChrW(8217) & "][!a-zA-Z]"
  .Replacement.Text = "\1 "
  .Execute Replace:=wdReplaceAll
  DoEvents

  .Text = "^p"
  .Replacement.Text = sp1
  .MatchWildcards = False
  .Execute Replace:=wdReplaceAll
  DoEvents
  
  .Text = "^t"
  .Execute Replace:=wdReplaceAll
  DoEvents

  .Text = "^p"
  .Execute Replace:=wdReplaceAll
  DoEvents
End With

' Create spelling error list
Set rng = ActiveDocument.Content

txt = rng.Text
' Ensure a single space between words
Do
  nowLen = Len(txt)
  txt = Replace(txt, sp2, sp1)
  DoEvents
Loop Until Len(txt) = nowLen

erList1 = ""
erListFreq1 = ""
erList2 = ""
erListFreq2 = ""

If Left(txt, 1) <> sp1 Then txt = sp1 & txt
Do
  spPos = InStr(Mid(txt, 2), " ")
  wd = Trim(Left(txt, spPos))
  spOK = Application.CheckSpelling(wd, MainDictionary:=langText)
  spEX = InStr(allExcepts, CR & wd & CR)
  nowLen = Len(txt)
  txt = Replace(txt, sp1 & wd & sp1, sp1)
  txt = Replace(txt, sp1 & wd & sp1, sp1)
  lossLen = Len(txt)
  If spOK = False And spEX = False And Len(wd) > 2 Then
  ' record it as a spelling error
    numErrs = Int((nowLen - lossLen) / Len(wd))
    numText = " . . . " & Trim(Str(numErrs))
    If LCase(wd) = wd Then
      erList2 = erList2 & wd & CR
      erListFreq2 = erListFreq2 & wd & numText & CR
    Else
      erList1 = erList1 & wd & CR
      erListFreq1 = erListFreq1 & wd & numText & CR
    End If
    StatusBar = wd & numText
    Debug.Print wd & numText
  End If
  DoEvents
Loop Until Len(txt) < 5

Selection.WholeStory
Selection.Delete
Selection.TypeText Replace(erList2, CR2, CR)
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric

Selection.EndKey Unit:=wdStory
Selection.TypeText CR
listStart = Selection.Start
Selection.TypeText CR & Replace(erList1, CR2, CR)
Selection.Start = listStart
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric

Selection.WholeStory
Selection.LanguageID = thisLanguage
Selection.Style = wdStyleNormal
Selection.Collapse wdCollapseStart
Selection.TypeText spellingListName & vbCr
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)

Documents.Add
Selection.TypeText Replace(erListFreq2, CR2, CR)
Selection.WholeStory
DoEvents
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
DoEvents
Selection.EndKey Unit:=wdStory
Selection.TypeText CR
listStart = Selection.Start
Selection.TypeText CR & Replace(erListFreq1, CR2, CR)
Selection.Start = listStart
DoEvents
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
DoEvents

Selection.WholeStory
Selection.LanguageID = thisLanguage
Selection.Style = wdStyleNormal
Selection.HomeKey Unit:=wdStory
Selection.TypeText spellingFreqName & vbCr
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)
Selection.Collapse wdCollapseStart

erList.Activate
Selection.HomeKey Unit:=wdStory
Application.ScreenUpdating = True
StatusBar = ""
If doingSeveralMacros = False Then
  totTime = Int(10 * (Timer - timeStart) / 60) / 10
  If totTime > 2 Then myResponse = MsgBox((totTime & "  minutes"), _
  vbOKOnly, "SpellAlyse")
  Beep
Else
  FUT.Activate
End If
Exit Sub

autoCorrect:
Application.ScreenUpdating = True
allText = FUT.Content.Text
If InStr(allText, "|") > 0 Then GoTo createListPair

myResponse = MsgBox("Auto-create FRedit list items?", _
     vbQuestion + vbYesNoCancel, "SpellAlyse")
If myResponse <> vbYes Then Exit Sub
Beep
myResponse = MsgBox("Turn track changes on?", _
     vbQuestion + vbYesNoCancel, "SpellAlyse")
If myResponse = vbCancel Then Exit Sub
myRev = (myResponse = vbYes)
If myRev Then ActiveDocument.TrackRevisions = True

' Add auto-created FRedit items
Set rng = ActiveDocument.Content

' (only first character, in case of split language)
rng.End = rng.Start + 1
langName = Languages(rng.LanguageID).NameLocal

i = 0
Do
  i = i + 1
  Set myPara = FUT.Paragraphs(i)
  wd = Replace(myPara.Range.Text, CR, "")
  If LCase(wd) = wd And wd <> "" Then
    DoEvents
    spellOK = Application.CheckSpelling(wd, _
         MainDictionary:=langName)
    Set suggList = Application.GetSpellingSuggestions(wd, _
         MainDictionary:=langName)
    newWord = ""
    DoEvents
    If suggList.count > 0 And Not (spellOK) Then
      newWord = suggList.Item(1).Name
      If myRev = True Then
        Set rng = myPara.Range.Duplicate
        rng.MoveEnd , -1
        myFR = "~<" & wd & ">|" & newWord
        rng.Text = myFR
      Else
        myPara.Range.Text = "~<" & wd & ">|" & newWord & CR
      End If
      myPara.Range.Select
      Selection.Collapse wdCollapseEnd
      DoEvents
    End If
  End If
  DoEvents
Loop Until i > 3 And myPara = CR
ActiveDocument.TrackRevisions = False
Exit Sub

createListPair:
' Make Flist + MarkIt list + Elist
Application.ScreenUpdating = True
myResponse = MsgBox("Create exceptions and FRedit lists?", _
     vbQuestion + vbYesNoCancel, "SpellAlyse")
If myResponse <> vbYes Then Exit Sub

Documents.Add
Set eList = ActiveDocument
Selection.TypeText Text:="| Elist" & CR2
eList.Paragraphs(1).Style = eList.Styles(wdStyleHeading1)
Set eRng = ActiveDocument.Content

Documents.Add
Set fList = ActiveDocument
Selection.TypeText Text:="| FRedit" & CR2
fList.Paragraphs(1).Style = fList.Styles(wdStyleHeading1)
Set fRng = ActiveDocument.Content

' FUT is now the spelling error list
For i = 1 To FUT.Paragraphs.count
  DoEvents
  Set itemRng = FUT.Paragraphs(i).Range
  If InStr(itemRng.Text, spellingListName) = 0 _
       And Len(itemRng.Text) > 3 Then
    DoEvents
    If InStr(itemRng.Text, "|") > 0 Then
      fRng.Collapse wdCollapseEnd
      fRng.FormattedText = itemRng.FormattedText
    Else
      myFontCol = itemRng.Font.Color
      myHiCol = itemRng.HighlightColorIndex
      If myFontCol > 0 Or myHiCol > 0 Then
        fRng.Collapse wdCollapseEnd
        fRng.Text = "~<" & Replace(itemRng.Text, CR, _
             "") & ">" & "|^&" & CR
        fRng.Expand wdParagraph
        fRng.Font.StrikeThrough = True
        If myFontCol > 0 Then fRng.Font.Color = myFontCol
        If myHiCol > 0 Then fRng.HighlightColorIndex = myHiCol
        fRng.Collapse wdCollapseEnd
      Else
        eRng.Collapse wdCollapseEnd
        eRng.Text = itemRng.Text
      End If
    End If
  End If
  DoEvents
Next i
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub

Sub SpellingErrorListerBilingual()
' Paul Beverley - Version 03.09.20
' Generates an alphabetic list all the bilingual spelling 'errors'

myLanguage_1 = wdEnglishUK
' myLanguage_2 = wdPortugueseBrazil
' myLanguage_2 = wdPortugueseBoth
myLanguage_2 = wdFrench

spellingListName = "SpellingErrors"


myFind = "´a,´e,¨a,¨e,¨o,¨u,ˆo,"
myReplace = "á,é,ä,ë,ö,ü,ô,"

CR = vbCr
CR2 = CR & CR

Set FUT = ActiveDocument
doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)

lang1 = Languages(myLanguage_1).NameLocal
lang2 = Languages(myLanguage_2).NameLocal

myLang = "Using " & lang1 & " and " & lang2 & " dictionary. OK?"
If doingSeveralMacros = False Then
  myResponse = MsgBox(myLang, vbQuestion + vbYesNoCancel, _
       "Spelling Error Lister")
  If myResponse <> vbYes Then Exit Sub
End If
timeStart = Timer

' lang = Languages(thisLanguage).NameLocal
Set rngOK = ActiveDocument.Content
OKstart = InStr(rngOK.Text, "OKwords")
If OKstart > 0 Then
  rngOK.Start = OKstart + 6
  OKwords = rngOK.Text
Else
  OKwords = ""
End If

' Change ligature characters into character pairs
myFind = myFind & "," & ChrW(-1280) & "," & ChrW(-1279) & _
     "," & ChrW(-1278) & "," & ChrW(-1277) & "," _
     & ChrW(-1276)
myReplace = myReplace & ",ff,fi,fl,ffi,ffl"
fnd = Split(myFind, ",")
rpl = Split(myReplace, ",")
For i = 0 To UBound(fnd)
  Set rng = ActiveDocument.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = fnd(i)
    .Wrap = wdFindContinue
    .Replacement.Text = rpl(i)
    .MatchCase = False
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With
  If ActiveDocument.Footnotes.Count > 0 Then
    Set rng = ActiveDocument.StoryRanges(wdFootnotesStory)
    With rng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = fnd(i)
      .Wrap = wdFindContinue
      .Replacement.Text = rpl(i)
      .MatchCase = False
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
  End If
  If ActiveDocument.Endnotes.Count > 0 Then
    Set rng = ActiveDocument.StoryRanges(wdEndnotesStory)
    With rng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = fnd(i)
      .Wrap = wdFindContinue
      .Replacement.Text = rpl(i)
      .MatchCase = False
      .MatchWildcards = False
      .Execute Replace:=wdReplaceAll
    End With
  End If
Next i

' Create spelling error list
erList1 = CR
erList2 = CR
numFootnotes = ActiveDocument.Footnotes.Count
numEndnotes = ActiveDocument.Endnotes.Count

myEnd = ActiveDocument.Content.End
For i = 1 To 3
  If myResponse = vbNo Then i = 3
  If i = 1 And numFootnotes = 0 Then i = 2
  If i = 1 Then Set rng = ActiveDocument.StoryRanges(wdFootnotesStory)
  If i = 2 And numEndnotes = 0 Then i = 3
  If i = 2 Then Set rng = ActiveDocument.StoryRanges(wdEndnotesStory)
  If i = 3 Then Set rng = ActiveDocument.Content
  For Each wd In rng.Words
    If Len(Trim(wd)) > 2 And LCase(wd) <> UCase(wd) And _
         wd.Font.StrikeThrough = False And wd <> "OKwords" Then
         padWd = " " & Trim(wd) & " "
      OKword = (InStr(OKwords, CR & Trim(wd) & CR) > 0)
      DoEvents
      If Application.CheckSpelling(wd, MainDictionary:=lang1) = False _
           And Application.CheckSpelling(wd, MainDictionary:=lang2) = _
           False And OKword = False Then
        pCent = Int((myEnd - wd.End) / myEnd * 100)

        ' Report progress
        If i = 1 Then myPrompt = "Checking footnote text."
        If i = 2 Then myPrompt = "Checking endnote text."
        If i = 3 Then myPrompt = "Checking main text."
        StatusBar = "Generating errors list. " & myPrompt & _
             " To go:  " & Trim(Str(pCent)) & "%"
        Debug.Print "Generating errors list. " & myPrompt & _
             " To go:  " & Trim(Str(pCent)) & "%"
        erWord = Trim(wd)
        lastChar = Right(erWord, 1)
        If lastChar = "'" Or lastChar = ChrW(8217) Then _
          erWord = Left(erWord, Len(erWord) - 1)
        myCap = Left(erWord, 1)
        If UCase(myCap) = myCap Then
          If InStr(erList2, CR & erWord & CR) = 0 Then _
              erList2 = erList2 & erWord & CR
        Else
          If InStr(erList1, CR & erWord & CR) = 0 Then _
              erList1 = erList1 & erWord & CR
        End If
      End If
    End If
  Next wd
Next i
mainFileName = ActiveDocument.Name
Documents.Add

Selection.TypeText erList1
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
erList1 = Selection
Selection.Delete

Selection.TypeText erList2
Selection.WholeStory
Selection.Sort SortOrder:=wdSortOrderAscending, _
     SortFieldType:=wdSortFieldAlphanumeric
erList2 = Selection
Selection.Delete

Selection.TypeText Text:=erList1
'Selection.TypeText CR2
Selection.TypeText Text:=erList2
'Selection.TypeText CR2

Selection.HomeKey Unit:=wdStory

If numFootnotes > 0 Then
  Selection.TypeText CR & "| footnotes = yes" & CR
End If
If numEndnotes > 0 Then
  Selection.TypeText CR & "| endnotes = yes" & CR
End If

StatusBar = ""
Selection.TypeText spellingListName & vbCr
ActiveDocument.Paragraphs(1).Style = ActiveDocument.Styles(wdStyleHeading1)

If doingSeveralMacros = False Then
  totTime = Int(10 * (Timer - timeStart) / 60) / 10
  If totTime > 2 Then myResponse = MsgBox((totTime & "  minutes"), _
  vbOKOnly, "Spelling Error Lister")
  Beep
Else
  FUT.Activate
End If
End Sub

Sub WordPairAlyse()
' Paul Beverley - Version 20.01.22
' Creates a file of all the adjacent word pairs

' Ignore these words
nonWords = "a,as"

myScreenOff = True

Set FUT = ActiveDocument

At = LCase(FUT.Content.Text)

doingSeveralMacros = (InStr(FUT.Name, "zzTestFile") > 0)
If doingSeveralMacros = False Then
  myResponse = MsgBox("      WordPairAlyse" & vbCr & vbCr & _
       "Find word pairs?", vbQuestion _
       + vbYesNoCancel, "WordPairAlyse")
  If myResponse <> vbYes Then Exit Sub
End If
If myScreenOff = True Then
  Application.ScreenUpdating = False
  On Error GoTo ReportIt
End If

startTime = Timer
chs = " , . ! : ; [ ] { } ( ) / \ + "
chs = chs & ChrW(8220) & " "
chs = chs & ChrW(8221) & " "
chs = chs & ChrW(8201) & " "
chs = chs & ChrW(8222) & " "
chs = chs & ChrW(8217) & " "
chs = chs & ChrW(8216) & " "
chs = chs & ChrW(8212) & " "
chs = chs & ChrW(8722) & " "
chs = chs & vbCr & " "
chs = chs & vbTab & " "
chs = " " & chs & " "
chs = Replace(chs, "  ", " ")
chs = Left(chs, Len(chs) - 1)

chars = Split(chs, " ")
For i = 1 To UBound(chars)
  At = Replace(At, chars(i), " " & chars(i) & " ")
Next i

' Remove all non-words
nonWords = "," & nonWords & ","
nonWords = Replace(nonWords, ",,", ",")
nonWords = Left(nonWords, Len(nonWords) - 1)

wd = Split(nonWords, ",")
Set rng = ActiveDocument.Content
For i = 1 To UBound(wd)
  At = Replace(At, " " & wd(i) & " ", " ")
  DoEvents
Next i
At = Replace(At, "  ", " ")

Documents.Add
Selection.Text = " " & At

Set rng = ActiveDocument.Content
Selection.HomeKey Unit:=wdStory

Set rng = ActiveDocument.Content
At = LCase(rng.Text)
myTot = Len(At)
If Asc(At) = 32 Then
  ptr = 2
Else
  ptr = 1
End If
ptrWas = ptr
Do
  ch = Mid(At, ptr, 1)
 ' Debug.Print ch & "|"
  ptr = ptr + 1
Loop Until ch = " "

w2 = Mid(At, ptrWas, ptr - ptrWas - 1)
' Debug.Print w2 & "|"

allChkd = " "
myResult = ""
Do
  w1 = w2
  ptrWas = ptr
  Do
    ch = Mid(At, ptr, 1)
    ptr = ptr + 1
  Loop Until ch = " "
  
  w2 = Mid(At, ptrWas, ptr - ptrWas - 1)
  
  If UCase(w1) <> w1 And UCase(w2) <> w2 Then
    oneWd = w1 & w2
    chk = " " & oneWd & " "
    If InStr(allChkd, chk) = 0 Then
      ' Check it!
      If InStr(At, chk) > 0 Then
        numSingle = Len(Replace(At, chk, chk & "!")) - myTot
        chk2 = " " & w1 & " " & w2 & " "
        numPair = Len(Replace(At, chk2, chk2 & "!")) - myTot
        myResult = myResult & w1 & " " & w2 & " . . " & _
             Trim(Str(numPair)) & vbCr
        myResult = myResult & oneWd & " . . " & _
             Trim(Str(numSingle)) & vbCr & vbCr
        Debug.Print Trim(Str(Int((myTot - ptr) / 6000))) _
             & ",000  to go" & "        " & w1 & " " & w2
        StatusBar = Trim(Str(Int((myTot - ptr) / 6000))) _
             & ",000  to go" & "        " & w1 & " " & w2
      End If
      allChkd = allChkd & oneWd & " "
    End If
  End If
  DoEvents
Loop Until InStr(Mid(At, ptr), " ") = 0

Selection.WholeStory
Selection.Delete
If Len(myResult) > 0 Then
  Selection.Text = myResult
  Set rng = ActiveDocument.Content
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "^p^p"
    .Replacement.Text = "zczc"
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  
    .Text = "^p"
    .Replacement.Text = ":"
    .Execute Replace:=wdReplaceAll
    
    .Text = "zczc"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll
  End With
  Set rng = ActiveDocument.Content
  rng.Sort SortOrder:=wdSortOrderAscending
  With rng.Find
    .Text = "^p"
    .Replacement.Text = "^p^p"
    .Execute Replace:=wdReplaceAll
    .Text = ":"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll
  End With
  
  Set rng = ActiveDocument.Content
  If Len(rng.Paragraphs(1)) < 3 Then rng.Paragraphs(1).Range.Delete
Else
  Selection.TypeText vbCr & "No word pairs found" & vbCr
End If
Selection.HomeKey Unit:=wdStory
Selection.TypeText "Word pair inconsistencies" & vbCr
ActiveDocument.Paragraphs(1).Style = _
     ActiveDocument.Styles(wdStyleHeading1)
timNow = Timer
timGone = timNow - startTime
m = Int(timGone / 60)
s = Int(timGone) - m * 60
Application.ScreenUpdating = True
If doingSeveralMacros = False Then
  Beep
  myTime = Timer
  Do
  Loop Until Timer > myTime + 0.3
  Beep
  myTime = Timer
  Do
  Loop Until Timer > myTime + 0.3
  Beep
  MsgBox "Total time:" & Str(m) & " m " & Str(s) & " s"
Else
  FUT.Activate
End If
Exit Sub

ReportIt:
Application.ScreenUpdating = True
On Error GoTo 0
Resume
End Sub

