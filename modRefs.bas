Attribute VB_Name = "modRefs"
Option Explicit

'TODO: Did declaring this external function violate macro secruity?
' Seems to be working now ???
' We're going to use a modal msgbox so declare it
Public Declare Function MsgBoxModal _
        Lib "User32" Alias "MessageBoxA" _
           (ByVal hWnd As Long, _
            ByVal lpText As String, _
            ByVal lpCaption As String, _
            ByVal wType As Long) _
        As Long
        
Dim strHeadings()   As String
Dim strNumbered()   As String

Sub update_all()
    Dim s As Range
    For Each s In ActiveDocument.StoryRanges
        s.Fields.Update
    Next
End Sub
Sub test()
    Debug.Print isLevel(2, Selection.Previous(wdWord))
End Sub
'TODO: Find out how much of the number format I can extract from Headings
'      [STYLE].ListLevels(?).NumberFormat/NumberStyle

Sub NumberFormatTest()
' An attempt to find the number formats based on the heading styles
    Dim s As Style
    Dim i As Integer
    Set s = ActiveDocument.Styles("Heading 1")
    Debug.Print s.NameLocal
    
    If Not s.ListTemplate Is Nothing Then
        With s.ListTemplate
            For i = 1 To .ListLevels.Count
                Debug.Print i, .ListLevels(i).NumberFormat, .ListLevels(i).NumberStyle
                .ListLevels(i).NumberStyle = wdListNumberStyleArabic1
            Next i
        End With
    End If
End Sub

Sub TestXref()
    '1. go through the field collection
    '2. is it an xRef?
    '3. unfieldify
    
    Dim f As Field
    Dim s As Range
    Dim t As String
        
    'Debug.Print ActiveDocument.Fields.Count
    'For Each s In ActiveDocument.StoryRanges
        For Each f In Selection.Fields
            'well is it an xref?
            If f.Type = wdFieldRef Then
                t = f.Result
                f.Section
                Selection.Fields.Update
                If f.Result <> t Then Stop
            End If
        Next f
    'Next s
    Debug.Print ActiveDocument.Fields.Count
End Sub


Sub UnFieldifyXRef()
    '1. go through the field collection
    '2. is it an xRef?
    '3. unfieldify
    
    Dim f As Field
    Dim s As Range
        
    Debug.Print ActiveDocument.Fields.Count
    For Each s In ActiveDocument.StoryRanges
        For Each f In s.Fields
            'well is it an xref?
            If f.Type = wdFieldRef Then
                'unfieldify
                f.Unlink
            End If
        Next f
    Next s
    Debug.Print ActiveDocument.Fields.Count
End Sub

Function IsDash(s As String) As Boolean
'TODO get around to adding all manner of dashes
    If s = "-" Then
        IsDash = True
    Else
        IsDash = False
    End If
End Function

Function IsValidConjunction(s As String) As Boolean
    Dim t   As String
    
    t = LCase(s)
    If t = "and" Or t = "or" Or t = "through" Or t = "to" Or IsDash(t) Then
        IsValidConjunction = True
    Else
        IsValidConjunction = False
    End If
End Function

Function IsAThing(s As String, blnClausesToo As Boolean) As Boolean
    Select Case LCase(Trimify(s))
        Case "section"
            IsAThing = True
        Case "sections"
            IsAThing = True
        Case "subsection"
            IsAThing = True
        Case "subsections"
            IsAThing = True
        Case "article"
            IsAThing = True
        Case "articles"
            IsAThing = True
        Case "paragraph"
            IsAThing = True
        Case "paragraphs"
            IsAThing = True
        Case Else
            IsAThing = isClauseType(s, blnClausesToo) And blnClausesToo
    End Select
End Function

Function isClauseType(str, blnClausesToo) As Boolean
    'TODO: include clauses as a thing until then punt
    isClauseType = False
    Exit Function
    
    Select Case (str)
        Case "clause"
            isClauseType = True And blnClausesToo
        Case "clauses"
            isClauseType = True And blnClausesToo
        Case "subclause"
            isClauseType = True And blnClausesToo
        Case "subclauses"
            isClauseType = True And blnClausesToo
        Case Else
            isClauseType = False
    End Select
End Function
Sub HyperLinkRef()
    '1. go through the field collection
    '2. is it an xRef?
    '3. hyperlink
Debug.Print Now
    Dim f As Field
    Dim s As Range
    Dim i As Integer
    Dim c As Long
        
    For Each s In ActiveDocument.StoryRanges
    '
        For Each f In s.Fields
            'well is it an xref?
            If f.Type = wdFieldRef Then
                'hyperlink
                ' if there isn't a \h in the field code
                ' put \h in from of the \w
                If InStr(f.Code, "\h") = 0 Then
                    i = InStr(f.Code, "\w")
                    ' if there isn't a \w in the field code punt
                    If i = 0 Then
                        Stop
                    Else
                        c = c + 1
                        f.Code.Select
                        f.Code.Text = Left(f.Code, i - 1) & "\h " & Mid(f.Code, i)
                        f.Update
                    End If
                End If
            End If
        Next f
    Next s
    
    Debug.Print c, Now
End Sub

Sub UnFieldifyXRefSelection()
    '1. go through the field collection
    '2. is it an xRef?
    '3. unfieldify
    
    Dim f As Field
    
    For Each f In Selection.Fields
        'well is it an xref?
        If f.Type = wdFieldRef Then
            'unfieldify
            f.Unlink
        End If
    Next f
    Debug.Print ActiveDocument.Fields.Count
End Sub

Function LCaseNoHardSpace(s As String) As String
    LCaseNoHardSpace = LCase(NoHardSpace(s))
End Function

Function NoHardSpace(s As String) As String
    NoHardSpace = Replace(s, Chr(160), " ")
End Function

Function Trimify(s As String) As String
'same as trim but also trim hard spaces
'3/30/2016 - are we still using this function?
'   Why not just trim(nohardspace(s))?
    s = Trim(s)
    If Len(s) Then
        Do While IsSpace(Mid(s, Len(s)))
            s = Trim(Left(s, Len(s) - 1))
        Loop
    End If
    Trimify = s
End Function

Sub AutoRef()
'TODO: Sections
'TODO: Section [X] or
'TODO: Section [X],
Debug.Print Now
    Dim w               As Range
    Dim s               As Range
    Dim iReply          As Long
    Dim fFieldShading   As Long
    Dim blnClausesToo   As Boolean
    
    blnClausesToo = True
    
    strHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    strNumbered = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
    
    fFieldShading = ActiveWindow.View.FieldShading
    'turn on show fieldcodes
    ActiveWindow.View.ShowFieldCodes = True
    
    Application.ScreenUpdating = False
    For Each s In ActiveDocument.StoryRanges
        For Each w In s.Words
            'TODO: Test if we're at the last word?
            If w.End = s.Words.Last.End Then Exit For
            
            If IsAThing(w.Text, blnClausesToo) And isNumericEx(w.Next(wdWord)) Or _
                (isInTheHole(NextNum(w)) And Not blnClausesToo) Then
                
                
            'TODO if we're not looking for clauses then if previous word is a clause punt
                If isClauseType(w.Text, blnClausesToo) Then MsgBox "we're NOT doing clauses"
            '  should the isinthehole text be something like inthehole and isclausetype != blnClauses too?
'            Stop
'Stop
            'If LCase(Trimify(w.Text)) = "sections" And _
              IsNumeric(w.Next(wdWord)) Then
                'turn off show field codes
                ActiveWindow.View.ShowFieldCodes = False
                
                w.Next(wdWord).Select
                
                ' this should handle clauses and subclauses
                If InStr(w, "clause") Then
                    Set w = w.Next(wdWord)
                    SelectWhole 'TODO:  Will this break it?
                End If
                
                Do While isNumericEx(Selection.Text) Or isInTheHole(Selection.Text)
                    ' TODO since we're doing the same thing for section/sections
                    ' , and prolly article/articles this should be broken out
                    ActiveWindow.View.ShowFieldCodes = False
                    Application.ScreenUpdating = True
                    Application.ScreenRefresh
                    
                    w.Select
                    SelectWhole
                    iReply = MsgBoxModal(&O0, "Cross reference to " & Selection.Text & "?", _
                        "Insert Cross Reference", vbYesNoCancel + vbSystemModal)
                    If iReply = vbYes Then
                        DoXRef
                    ElseIf iReply = vbCancel Then
                        ActiveWindow.View.ShowFieldCodes = True
                        Exit For
                    End If
                    
                    Selection.Next(wdWord).Select
                    Application.ScreenUpdating = False
                    ActiveWindow.View.ShowFieldCodes = True
                    'TODO test for and/or case 'TODO seems to work for and but not for or?
                    If IsValidConjunction(Trim(Selection.Text)) Then
                        Selection.Next(wdWord).Select
                        'TODO: if the selection is a "(" then extend it to the next ")"
                        'TODO: commented out because sometimes it broke the isnumericex test :-(
                        ''SelectWhole 'TODO will this do the above or break everything?
                    End If
                    Set w = Selection.Range
                Loop
            End If
            ActiveWindow.View.ShowFieldCodes = True 'TODO put this where we need it?
        Next w
    Next s
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    ActiveWindow.View.ShowFieldCodes = False
    MsgBox "ALL Done!"
    ActiveWindow.View.FieldShading = fFieldShading
    ' unload strheadings
    Erase strNumbered
    Erase strHeadings
Debug.Print Now
End Sub

Sub xRef()
    strHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    strNumbered = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
    DoXRef
    Erase strHeadings
    Erase strNumbered
End Sub

Sub DoXRef()
    ' Assumes that if L1 = Article # then L2 = Section #
    '   Hmmm? Does it still do this?
    ' popping up and down levels kinds of screams 4 recursion not
    ' iteration so we're going give it a go
    
    Dim strS            As String
    Dim i               As Long
    Dim strParaNum      As String
    
    ' Select whole now needs strheadings to check for "Section" in the numbers :-(
    ' TODO: Make selectwhole elegant again :-(
    SelectWhole
    
    If isInTheHole(Selection.Text) Then
        ' This is for reference to plain paragraph #'s "in the hole"
        ' prompt for paragraph #
        strS = Selection.Text
        LoadDropDown "", strS, 1, 1
        With frmMine
            If .lstBox.ListCount > 0 Then
                .lstBox.ListIndex = 0
                .lblQuery = "Which " & strS & " are we talking about?"
                .Show
                strS = .lstBox.Text
                i = .lstBox.List(.lstBox.ListIndex, 1)
            End If
        End With
        Unload frmMine
        If strS = "" Then
            Exit Sub
        End If
    Else
        strS = LCaseNoHardSpace(Selection.Text)
    End If
    
    If i = 0 Then
        i = SearchHeadings("", strS, 1, 1)
    Else
        insertXRef wdRefTypeNumberedItem, i, True
    End If
    
    ' back up and select the whole thing
    Selection.MoveLeft
    SelectWhole
    
    ' if we haven't found it but it's an article try just checking the number
    'If Not SelectionInAField And IsArticle(Selection.Text) Then
    ' Changed that to just a space
    If InStr(strS, " ") > 1 Then
        SelectNumberOnly
        i = SearchHeadings("", Selection.Text, 1, 1)
    End If

    SelectWhole
    
    ' if we still haven't found it then try looking through the numbered items
    If Not SelectionInAField Then
        SelectWhole
        i = SearchNumberedItems("", strS, 1, 1)
        Selection.Collapse wdCollapseEnd
    End If
    
    SelectWhole
    ' if we STILL haven't found it but it's an article try just checking the number
    If Not SelectionInAField And IsArticle(Selection.Text) Then
        SelectNumberOnly
        'TRY AGAIN
        i = SearchHeadings("", Selection.Text, 1, 1)
    End If
    
    'back up one, just do it
    Selection.MoveLeft
    SelectWhole
    
    ' if we haven't found anything make selection red & doubleunderline
    ' so we can easily find it (assumes nothing else red & dbl UL in doc)
    If Not SelectionInAField Then
        With Selection
            .Font.Color = wdColorRed
            .Font.Underline = wdUnderlineDouble
        End With
    End If
    
    
    Selection.Collapse wdCollapseEnd
End Sub

Sub SelectNumberOnly()
    Selection.Words(2).Select
    Selection.Previous(wdCharacter).Text = Chr(160)
    Selection.MoveRight
    SelectWhole
End Sub
Function getIndent(str As String) As Integer
' indent is the position of the first character that's not a space
    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) <> " " Then
            getIndent = i
            Exit Function
        End If
    Next i
End Function

Function SearchHeadings(strStub As String, strLookfor As String, index As Long, indent As Integer) As Long
    Dim strCurrH        As String
    Dim strLevelNum     As String
    Dim i               As Long
    Dim j               As Long
    Dim iNewIndent      As Integer
    Dim iOldIndent      As Integer
    Dim isNumberOnly    As Boolean
    
    For i = index To UBound(strHeadings)
        iNewIndent = getIndent(strHeadings(i))
        If i > 1 Then
            iOldIndent = getIndent(strHeadings(i - 1))
        End If
        
        Do While indent <> iNewIndent
            If indent < iNewIndent Then
                ' everytime indent increases call this function with i as index & add 2 to indent
                ' return value of searchy as index to this loop
                i = SearchHeadings(strStub + strLevelNum, strLookfor, i, indent + 2)
                'If i = -1 Then
                If i <= 0 Then
                    SearchHeadings = -1
                    Exit Function
                End If
                iNewIndent = getIndent(strHeadings(i))
            Else
                'everytime indent decreases return index
                SearchHeadings = i
                Exit Function
            End If
        Loop
        
        If i > UBound(strHeadings) Or i = 0 Then
            SearchHeadings = -1
            Exit Function
        End If
            
        strCurrH = LCaseNoHardSpace(Trim(strHeadings(i)))
        
        ' is this an article or section?
        If isLevel(1, strCurrH) Then
            strStub = ""
            strLevelNum = GetStub(strCurrH)
        ElseIf isLevel(2, strCurrH) Then
            strStub = GetStub(strCurrH)
            strLevelNum = ""
        Else
            strLevelNum = Trim(Left(strCurrH, InStr(strCurrH, " ")))
            If Len(strLevelNum) Then
                If Mid(strLevelNum, Len(strLevelNum)) = "." Then
                    strLevelNum = Left(strLevelNum, Len(strLevelNum) - 1)
                End If
            End If
        End If
                
        'if found return -1
        If InStr(strStub & strLevelNum, strLookfor) Then
            If Len(strLookfor) <> Len(strStub & strLevelNum) Then
                isNumberOnly = True
            End If
            If insertXRef(wdRefTypeHeading, i, isNumberOnly) Then
                FindTheRest strLookfor, wdRefTypeHeading, i, isNumberOnly
                SearchHeadings = -1
                Exit Function
            End If
        End If
    Next i
End Function

Sub FindTheRest(strLookfor As String, iRefType As Long, index As Long, numberOnly As Boolean)
    Dim iReply              As Long
    Dim rngOrignalPosition  As Range
    Dim rngFirstPosition    As Range
    Dim strOldSearch        As String
    Dim strOldReplace       As String
    Dim blnFirstTime        As Boolean
    Dim blnOldMatchWild     As Boolean
    Dim lngOldWrap          As Long
    Dim blnOldFormat        As Boolean
    Dim blnOldMatchCase     As Boolean
    Dim blnOldForward       As Boolean
    Dim blnOldMatchByte     As Boolean
    Dim blnOldMatchSound    As Boolean
    Dim blnOldMatchForms    As Boolean
    
    'TODO: Let's skip doing this for a while
    Exit Sub
    
    If isInTheHole(Selection.Fields(1).Result) Then _
        Exit Sub
    
    'If MsgBox("Search for other occurences of " & strLookfor, vbYesNo) = vbNo Then
    '    Exit Sub
    'End If
    
    Set rngOrignalPosition = Selection.Range
    blnFirstTime = True
    
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        ' save the commoner old find parameters
        strOldSearch = .Text
        strOldReplace = .Replacement.Text
        blnOldMatchWild = .MatchWildcards
        lngOldWrap = .Wrap
        blnOldFormat = .Format
        blnOldMatchCase = .MatchCase
        blnOldForward = .Forward
        blnOldMatchByte = .MatchByte
        blnOldMatchSound = .MatchSoundsLike
        blnOldMatchForms = .MatchAllWordForms
        
        .Text = strLookfor
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue  'Is this what we need for our purposes
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Do While Selection.find.Execute
        'If it's already a field or the next character is "(" then don 't even bother to ask
        If (Not SelectionInAField _
                And Selection.Next(wdCharacter).Text <> "(" _
                And Not IsNumeric(Selection.Next(wdCharacter)) _
                And Not IsNumeric(Selection.Previous(wdCharacter))) _
                And Not IsUpper(Selection.Next(wdCharacter)) _
                And Not (Selection.Next(wdCharacter) = "." And _
                IsNumeric(Selection.Next(wdCharacter).Next(wdCharacter))) _
                And Not (Selection.Previous(wdCharacter) = "." And _
                IsNumeric(Selection.Previous(wdCharacter).Previous(wdCharacter))) _
                Then
            
            If blnFirstTime Then
                Set rngFirstPosition = Selection.Range
                rngOrignalPosition.Select
                If MsgBox("Cross reference other occurences of " _
                  & strLookfor, vbYesNo) = vbNo Then
                    Exit Do
                Else
                    blnFirstTime = False
                    rngFirstPosition.Select
                End If
            End If
            
            If LCaseNoHardSpace(Selection.Text) = strLookfor Then
                iReply = MsgBoxModal(&O0, _
                  "Insert Cross Reference Here?", "My Question", vbYesNoCancel + vbSystemModal)
                'TODO: see if a modal messagebox works
                '''iReply = MsgBox("Insert Cross Reference Here?", vbYesNoCancel)
                If iReply = vbCancel Then
                    Exit Do
                ElseIf iReply = vbYes Then
                    insertXRef iRefType, index, numberOnly
                End If
            End If
        End If
    Loop
    
    'go back to where we were
    With Selection.find
        .Text = strOldSearch
        .Replacement.Text = strOldReplace
        '.MatchWildcards = blnOldMatchWild  'TODO: Find out how to make match wild card stick!
        .Wrap = lngOldWrap
        .Format = blnOldFormat
        .MatchCase = blnOldMatchCase
        .Forward = blnOldForward
        .MatchByte = blnOldMatchByte
        .MatchSoundsLike = blnOldMatchSound
        .MatchAllWordForms = blnOldMatchForms
    End With
    
    rngOrignalPosition.Select
End Sub

Function insertXRef(iRefType As Long, _
  index As Long, _
  onlyNumber As Boolean) As Boolean
    Dim flgUL       As Boolean
    Dim flgBold     As Boolean
    Dim flgItal     As Boolean
    Dim iOldStart   As Long
    Dim iRefKind    As Long
    Dim strOld      As String
    Dim strNew      As String
    Application.ScreenUpdating = False
    With Selection
        ' save the font formatting
        ' formatting doesn't seem to work for level 2-
        iOldStart = .Start
        With .Characters.First
            flgUL = .Underline
            flgBold = .Bold
            flgItal = .Italic
        End With
        
        strOld = Selection.Text
        
        If isInTheHole(Selection.Text) Then
            iRefKind = wdNumberNoContext
        Else
            iRefKind = wdNumberFullContext
        End If
        
        .InsertCrossReference _
            ReferenceType:=iRefType, ReferenceKind:=iRefKind, _
            ReferenceItem:=index, InsertAsHyperlink:=True
         
        ' just back up 1 before trying to select the whole field
        Selection.MoveLeft wdCharacter
        SelectWhole
        strNew = Selection.Text
        
        If onlyNumber Then
            ' add the \t to the field code
            Application.ScreenUpdating = True
            SelectWhole 'this don't work no more
            Selection.Words(1).Select
            
            With Selection.Fields(1)
                .Code.Text = .Code.Text & "\t"
                .Update
            End With
            Application.ScreenUpdating = False
        End If
        
        If strOld <> Selection.Text Then
            On Error Resume Next
            If strOld = Selection.Fields(1).Result Then
                insertXRef = True
            ElseIf isLevel(1, LCase(strOld)) Or isLevel(2, LCase(strOld)) Then
                If LCaseNoHardSpace(strOld) = LCaseNoHardSpace(Selection.Text) Then
                    insertXRef = True
                End If
            Else
                Selection.Text = strOld
                insertXRef = False
            End If
            On Error GoTo 0
        Else
            insertXRef = True
        End If
            
        ' match font formatting! UL/Bold?Ital
        If .Start = .End Then
            Selection.MoveLeft wdCharacter
        End If
        
        SelectWhole '// doesn't work with tables OMG!
        With Selection.Range
            .Bold = flgBold
            .Underline = flgUL
            .Italic = flgItal
        End With
    End With
    Application.ScreenUpdating = True
End Function

Function SearchNumberedItems(strStub As String, _
                                strLookfor As String, _
                                index As Long, _
                                indent As Integer) As Long
    Dim strCurrH        As String
    Dim strLevelNum     As String
    Dim i               As Long
    Dim j               As Long
    Dim iSpaceAt        As Integer
    Dim iNewIndent      As Integer
    Dim isNumberOnly    As Boolean

    For i = index To UBound(strNumbered)
        ' indent is the position of the first character that's not a Space
        iNewIndent = getIndent(strNumbered(i))
        'TOQ - DO WE NEED TO POP UP ON THE WHOLE OLDINDENT/NEWINDENT THING :-(

        
        If indent <> iNewIndent Then
            If indent < iNewIndent Then
                ' everytime indent increases [or we hit a section level]
                ' call this function with i as index & add 2 to indent
                ' return value of searchy as index to this loop

''which we haven't started yet
''do you remember why we used indent+2 instead of newindent?
''that's the sort of thing that needed a comment LOL
                'i = SearchNumberedItems(strStub + strLevelNum, strLookfor, i, indent + 2)
                i = SearchNumberedItems(strStub + strLevelNum, strLookfor, i, iNewIndent)
                If i = -1 Then
                    SearchNumberedItems = -1
                    Exit Function
                End If
            Else
                'everytime indent decreases return index
                SearchNumberedItems = i
                Exit Function
            End If
        End If
        
        If i > UBound(strNumbered) Or i = 0 Then
            SearchNumberedItems = -1
            Exit Function
        End If
        
        strCurrH = LCaseNoHardSpace(Trim(strNumbered(i)))
        ' is this an article or section?
        If isLevel(1, strCurrH) Or isLevel(2, strCurrH) Then
            ' works cause "article" and "section" are both 7 characters long
            strStub = GetStub(strCurrH)
            strLevelNum = ""
        Else
            'if no space in strcurrh just use the whole damn thing!
            iSpaceAt = InStr(strCurrH, " ")
            If iSpaceAt > 0 Then
                strLevelNum = Trim(Left(strCurrH, iSpaceAt))
            Else
                strLevelNum = strCurrH
            End If
        End If
        
        'if found insert xRef & return -1
        If InStr(strStub & strLevelNum, strLookfor) Then
            If Len(strLookfor) <> Len(strStub & strLevelNum) Then
                isNumberOnly = True
            End If
            If insertXRef(wdRefTypeNumberedItem, i, isNumberOnly) Then
                FindTheRest strLookfor, wdRefTypeNumberedItem, i, isNumberOnly
                SearchNumberedItems = -1
                Exit Function
            End If
        End If
    Next i
End Function

Function GetStub(str As String) As String
    'TODO: If this fails revert to GetStub OLD
    Dim strC    As String
    Dim strL    As String
    Dim i       As Integer
    
    'Stop
    If isLevel(1, str) Or isLevel(2, str) Then
        strL = getFirstLevelLabel
        i = InStr(Len(strL) + 2, str, " ")
        If i > 0 Then
            strC = Left(str, i - 1)
        End If
    Else
        strC = str
    End If
    
    'BUT if the last character is a "." we do want to cut it off
    If Right(strC, 1) = "." Then
        strC = Left(strC, Len(strC) - 1)
    End If
    GetStub = strC
End Function

Function GetStubOLD(str As String) As String
    Dim strC    As String
    Dim i       As Integer
    
    i = InStr(Len("article ") + 1, str, " ")
    
    If i > 0 Then
        strC = Left(str, i - 1)
    Else
        strC = str
    End If
    'BUT if the last character is a "." we do want to cut it off
    If Right(strC, 1) = "." Then
        strC = Left(strC, Len(strC) - 1)
    End If
    GetStubOLD = strC
End Function

Function IsArticle(s As String) As Boolean
    'IsArticle = InStr(LCase(s), "article")
    IsArticle = LCase(s) Like "article*"
End Function

Function IsSection(s As String) As Boolean
    'TODO: Delete this we don't call it anymore
    'TODO: We should also stop calling IsArticle
    'IsSection = InStr(LCase(s), "section")
    IsSection = LCase(s) Like "section*"
End Function

Function isLevel(level As Integer, str As String) As Boolean
    Dim s       As Style
    Dim strL    As String
    Dim i       As Integer
    
    Set s = ActiveDocument.Styles("Heading " & level)
    If Not s.ListTemplate Is Nothing Then
        strL = LCaseNoHardSpace(s.ListTemplate.ListLevels(level).NumberFormat)
        If IsAlpha(Left(strL, 1)) Then
            i = InStr(strL, "%")
            If LCase(str) Like Trim(Left(strL, i - 1)) & "*" Then
                isLevel = True
            End If
        End If
    End If
End Function

Function getFirstLevelLabel() As String
    Dim s As Style
    Dim l As String
    Dim i As Integer
    Set s = ActiveDocument.Styles("Heading 1")
    If Not s.ListTemplate Is Nothing Then
'        Stop
        l = s.ListTemplate.ListLevels(1).NumberFormat
        i = InStr(l, "%")
        getFirstLevelLabel = Trim(Left(l, i - 1))
    End If
    
End Function

Function IsUpper(c As String) As Boolean
    IsUpper = c >= "A" And c <= "Z"
End Function

Function IsLower(c As String) As Boolean
    IsLower = c >= "a" And c <= "z"
End Function

Function IsAlpha(c As String) As Boolean
    IsAlpha = IsUpper(c) Or IsLower(c)
End Function

Function IsParen(c As String) As Boolean
    IsParen = c = "(" Or c = ")"
End Function

Function IsLegal(c As String) As Boolean
    IsLegal = IsUpper(c) Or IsLower(c) Or IsNumeric(c) _
            Or IsParen(c) Or c = "."
End Function

Function IsSpace(c As String) As Boolean
    IsSpace = c = " " Or c = Chr(160)
End Function

Function isNumericEx(s As String) As Boolean
    isNumericEx = IsNumeric(s) Or isRomanNumeral(s)
End Function

Function isRomanNumeral(s As String) As Boolean

    Dim i As Integer
    Dim c As String
    
    If s = " " Then
        isRomanNumeral = False
    Else
        
        For i = 1 To Len(Trim(s))
            c = LCase(Mid(s, i, 1))
            If c <> "i" And c <> "v" And c <> "x" _
              And c <> "l" And c <> "c" Then
                isRomanNumeral = False
                Exit Function
            End If
        Next i
        isRomanNumeral = True
    End If
                
End Function

Function NextNum(r As Range) As String
'keep adding word while the last character of the word
' isaletter or isnumeric or isparen
'TODO: Check for balanced parens on close paren!!!
    Dim rNext   As Range
    Dim s       As String
    Dim c       As String
    Dim i       As Integer
  
'Stop
    Set rNext = r.Next(wdWord)
    Do While Not rNext Is Nothing
        c = rNext.Characters.Last
        s = s & rNext.Text
        If Not IsAlpha(c) And Not IsNumeric(c) And Not IsParen(c) Then
            Exit Do
        End If
        Set rNext = rNext.Next(wdWord)
    Loop
    NextNum = Trim(s)
End Function

Function isInTheHole(str As String) As Boolean
    'is it one characters in parens or a ruman numeral inparens?
    Dim strNumeral  As String
    
    If Len(str) = 3 Then
        isInTheHole = Mid(str, 1, 1) = "(" And Mid(str, 3, 1) = ")"
    ElseIf Len(str) > 1 Then
        strNumeral = LCase(Mid(str, 2, Len(str) - 2))
        If isRomanNumeral(strNumeral) Then
            isInTheHole = Mid(str, 1, 1) = "(" And Mid(str, Len(str), 1) = ")"
        Else
            isInTheHole = False
        End If
    End If
End Function

Function isXInTheHole(str As String) As Boolean
'TODO: 1-see if below def makes sense, 2-code it
'first anything in parens = isXInTheHole
'second anything in parens followed by alpha, paren or number = isXInTheHole
End Function

Function SelectionInAField() As Boolean
    'TODO: write a foolprof check to see if the selection is in a field
    ' for now if the (1) fields.count proeprty is > 1 or
    ' (2)the style has the word TOC in it then we assume we're in a field
    ' you'd think (1) would be sufficient but it aint punts for TOC
    ' obviously (2) will fail for TOCs if TOC styles are not used for the TOC
    ' & will generate a false positive if TOC Styles are used outside of the TOC
    ' but TT such is life
    With Selection
        If .Fields.Count > 0 Then
            SelectionInAField = True
        ElseIf .Words(1).Fields.Count > 0 Then
            SelectionInAField = True
        ElseIf .Style.NameLocal Like "TOC*" Then
            SelectionInAField = True
        Else
            SelectionInAField = False
        End If
    End With
End Function

Function BalancedParens(s As String) As Boolean
    Dim iLeft   As Integer
    Dim iRight  As Integer
    Dim i       As Integer
    
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "(" Then
            iLeft = iLeft + 1
        ElseIf Mid(s, i, 1) = ")" Then
            iRight = iRight + 1
        End If
    Next
    BalancedParens = iLeft = iRight
End Function

Sub SelectWhole()
'TODO: NOW it doesn't work when we're at the end of a paragraph, did it ever
'TODO: Restrict it to 2 words
    Dim i           As Integer
    Dim c           As Range
    Dim s           As String
    Dim iEnd        As Long
    Dim iStart      As Long
    Dim isEndOfCell As Boolean
'Stop
    iEnd = ActiveDocument.Range.End
    iStart = ActiveDocument.Range.Start
    
    ' collapse the selection
    Selection.Collapse wdCollapseStart
    
    ' go left until you hit a space or the beginning of the document
    If Selection.Start <> iStart Then
        Set c = Selection.Previous(wdCharacter)
        Do While IsLegal(c.Text) And c.Start <> iStart
            Selection.MoveLeft Extend:=True
            Set c = Selection.Characters.First
        Loop
        'TODO: see if this fixes anything
        'TODO: I think we're testing for too many exceptions :-(
        If Selection.Start <> iStart Then
            If IsSpace(Selection.Characters.First) Then
                'TODO: figure out why we didn't stop at a space going right
                Selection.MoveRight Extend:=True
            ElseIf isLevel(1, Selection.Previous(wdWord)) Or _
              isLevel(2, Selection.Previous(wdWord)) Then
                Selection.MoveLeft (wdWord)
            Else
                Selection.MoveRight Extend:=True
            End If
        End If
    End If

    iStart = Selection.Start
    ' go right until you hit something that's not a letter, number or paren
    ' or space with a number after it
    Set c = Selection.Characters.Last
    Do While IsLegal(c.Text) Or IsSpace(c.Text) And c.End <> iEnd
        If c.Text = "." Or IsSpace(c.Text) Then
            'TODO: Now we're looking at the next word instead of character.
            '       hopefully this doesn't break anything
            'If Not isNumericEx(c.Next(wdCharacter)) Or
            If Not isNumericEx(c.Next(wdWord)) Or _
              (c.Text = " " And Selection.Words.Count > 1) Then
                 Exit Do
            End If
        End If
        'if we're at the end of cell marker punt
        If AscW(Selection.Next(wdCharacter)) = 13 And Selection.Next(wdCharacter) <> vbCr Then
            isEndOfCell = True
            Exit Do
        End If
        Selection.End = Selection.End + 1
        Set c = Selection.Characters.Last
    Loop
    Selection.Start = iStart
    'if we're NOT at the end of cell character go back!
    If Not isEndOfCell Then ' AscW(Selection.Next(wdCharacter)) <> 13 Then 'Or Selection.Next(wdCharacter) = vbCr Then
        Selection.End = Selection.End - 1
    End If
    
    If Selection.Characters.First = "(" And Not isInTheHole(Selection.Text) Then
        If Selection.Characters.Last.Text = ")" Then
            If Not BalancedParens(Selection.Text) Then
                Selection.End = Selection.End - 1
            End If
        End If
        ' if still not in the hole punt
        If Not isInTheHole(Selection.Text) Then
            Selection.Start = Selection.Start + 1
        End If
    End If
    
    If Selection.Characters.Last.Text = ")" Then
        If Not BalancedParens(Selection.Text) Then
            Selection.End = Selection.End - 1
        End If
    End If
    
    'if selection starts with "section" and none of the headers have section shrink it!
    'TODO: What about articles? or Sections with "Section" headings
    If LCase(Selection.Text) Like "sections*" Or LCase(Selection.Text) Like "articles*" Then
        Selection.Start = Selection.Start + Len("section* ")
    ElseIf LCase(Selection.Text) Like "section*" Then
        'if strheadings not set then no need to check
        If IsArrayAllocated(strHeadings) Then
            For i = 1 To UBound(strHeadings)
                If LCaseNoHardSpace(Trim(strHeadings(i))) Like "section *" Then
                    Exit For
                End If
            Next i
            If i > UBound(strHeadings) Then
                Selection.Start = Selection.Start + Len("section ")
            End If
        End If
    End If
End Sub

Function IsArrayAllocated(Arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(Arr) And _
                           Not IsError(LBound(Arr, 1)) And _
                           LBound(Arr, 1) <= UBound(Arr, 1)
End Function

Function LoadDropDown(strStub As String, strLookfor As String, _
    index As Integer, indent As Integer) As Long
 'Based on SearchNumberedItems - so we can probably merge them
' TODO: pull out the bits of this that correspond to SearchNumberedItems

    Dim strCurrH        As String
    Dim strLevelNum     As String
    Dim i               As Integer
    Dim j               As Integer
    Dim iSpaceAt        As Integer
    Dim iNewIndent      As Integer
    Dim isNumberOnly    As Boolean

    For i = index To UBound(strNumbered)
        ' indent is the position of the first character that's not a Space
        iNewIndent = getIndent(strNumbered(i))
        
        If indent <> iNewIndent Then
            If indent < iNewIndent Then
                ' everytime indent increases [or we hit a section level]
                ' call this function with i as index & add 2 to indent
                ' return value of searchy as index to this loop
                i = LoadDropDown(strStub + strLevelNum, strLookfor, i, iNewIndent)
            Else
                'everytime indent decreases pop back to caller
                LoadDropDown = i
                Exit Function
            End If
        End If
        
        If i > UBound(strNumbered) Or i = 0 Then
            ' when we hit the end of the array pop back to caller
            Exit Function
        End If
        
        'TODO: Maybe this needs to be Original Case
        'TODO: Lcase only "article" & "section"
        'TODO: Just for load dropdowns? Other cases "seems" to work
        'strCurrH = LCaseNoHardSpace(Trim(strHeadings(i)))
        strCurrH = NoHardSpace(Trim(strNumbered(i)))
        
        ' is this an article or section?
        If isLevel(2, LCase(strCurrH)) Then
            strStub = GetStub(strCurrH)
            strLevelNum = ""
        ElseIf isLevel(1, LCase(strCurrH)) Then
            strStub = ""
            strLevelNum = ""
        Else
            'if no space in strcurrh just use the whole damn thing!
            iSpaceAt = InStr(strCurrH, " ")
            If iSpaceAt > 0 Then
                strLevelNum = Trim(Left(strCurrH, iSpaceAt))
            Else
                strLevelNum = strCurrH
            End If
        End If
        
        'if found add to the lstbox and carry on
        If strStub & strLevelNum Like "*" & strLookfor Then
            With frmMine
                .lstBox.AddItem strStub & strLevelNum
                .lstBox.List(.lstBox.ListCount - 1, 1) = i
            End With
        End If
    Next i
End Function

'TODO: Maybe maybe not
Function extendedIsInTheHole(str As String) As Boolean
' make sure everything in parens is a single character, a number
' or a roman numeral
    Dim i As Long
    Dim j As Long

    extendedIsInTheHole = True
    If isInTheHole(str) Then
        extendedIsInTheHole = True
    Else
        i = 1
        Do While i <> 0
            If Mid(str, i, 1) = "(" Then
                j = InStr(i + 1, str, ")")
                If j = 0 Then
                    extendedIsInTheHole = False
                    Exit Do
                ElseIf Not isInTheHole(Mid(str, i, j)) Then
                    extendedIsInTheHole = False
                    Exit Do
                End If
            End If
            i = InStr(i + 1, str, "(")
        Loop
    End If
End Function

Sub upp()
    Dim s As Range
    
    For Each s In ActiveDocument.StoryRanges
        s.Fields.Update
    Next
End Sub

Sub findref()
    Dim objfld As Field
    For Each objfld In ActiveDocument.Fields
        ' If the field is a cross-ref, do something to it.
        If objfld.Type = wdFieldRef Then
            objfld.Select
            Stop
        End If
    Next
End Sub


Sub testit()
    Dim x
    Dim y
    
    x = "dog"
    y = "ninny"
    Select Case x
        Case "dog"
        Case "cow"
        Case "moose"
        Case "cat"
            y = "how now"
        Case "kitten"
            y = "how then"
    End Select
    Debug.Print y
    
    
End Sub

Sub givemesometabs()
    Dim c   As Cell
    Dim r   As Range
    
    Set r = Selection.Range
    
    For Each c In r.Cells
        With c.Range.Characters
            .Item(.Count - 1).InsertAfter vbTab
        End With
    Next
End Sub
