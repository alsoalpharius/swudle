VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   12420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15540
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardChosen As Integer
Dim cardGuessed As String
Dim cardGuessedNo As Integer
Dim GuessNo As Integer
Dim isTitle As Integer
Dim TitleSplits() As String
Dim ATSplitGuess() As String
Dim ATSplitChosen() As String
Dim matchFound As Boolean
Dim splitGuessLength As Integer
Dim splitChosenlength As Integer
Dim ATGuess As String
Dim ATChosen As String

Private Sub CommandButton1_Click()
    cardGuessed = LCase(GuessBox.Text)
    
    isTitle = InStr(1, cardGuessed, ",")
    If (isTitle = 0) Then 'there is no title
        For x = 2 To 262
            If (LCase(Cells(x, 1)) = cardGuessed) Then
                cardGuessedNo = x
                Exit For
            End If
        Next
    
    Else 'there is a title
        TitleSplits = Split(cardGuessed, ", ")
        For x = 2 To 262
            If (LCase(Cells(x, 1)) = TitleSplits(0) And LCase(Cells(x, 2)) = TitleSplits(1)) Then
                cardGuessedNo = x
                Exit For
            End If
        Next
    End If
    
    If x = 263 Then
        MsgBox ("There is no card with that name")
        Exit Sub
    End If
    
    matchFound = False
       
    Select Case GuessNo
    Case 1
        Row11.Caption = Cells(cardGuessedNo, 1)
        Row12.Caption = Cells(cardGuessedNo, 2)
        Row13.Caption = Cells(cardGuessedNo, 3)
        Row14.Caption = Cells(cardGuessedNo, 4)
        Row15.Caption = Cells(cardGuessedNo, 5)
        Row16.Caption = Cells(cardGuessedNo, 6)
        Row17.Caption = Cells(cardGuessedNo, 7)
        Row18.Caption = Cells(cardGuessedNo, 8)
        Row19.Caption = Cells(cardGuessedNo, 9)
        Row110.Caption = Cells(cardGuessedNo, 10)
        
        For x = 1 To 10
            matchFound = False
            If (x = 5 Or x = 9) Then
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row1" & x).BackColor = vbGreen
                Else
                
                    isTitle = InStr(1, Cells(cardGuessedNo, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATGuess = Cells(cardGuessedNo, x)
                       splitGuessLength = 0
                    Else
                        ATSplitGuess = Split(Cells(cardGuessedNo, x), ", ")
                        splitGuessLength = UBound(ATSplitGuess) - LBound(ATSplitGuess)
                    End If
                    
                    isTitle = InStr(1, Cells(cardChosen, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATChosen = Cells(cardChosen, x)
                       splitChosenlength = 0
                    Else
                        ATSplitChosen = Split(Cells(cardChosen, x), ", ")
                        splitChosenlength = UBound(ATSplitChosen) - LBound(ATSplitChosen)
                    End If
                    
                    If (splitGuessLength > 0 And splitChosenlength > 0) Then
                        For y = 0 To splitGuessLength
                            For Z = 0 To splitChosenlength
                                If ATSplitGuess(y) = ATSplitChosen(Z) Then
                                    matchFound = True
                                End If
                            Next
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength > 0) Then
                        For Z = 0 To splitChosenlength
                            If ATGuess = ATSplitChosen(Z) Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength > 0 And splitChosenlength = 0) Then
                        For y = 0 To splitGuessLength
                            If ATSplitGuess(y) = ATChosen Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength = 0) Then
                        If ATGuess = ATChosen Then
                            matchFound = True
                        End If
                    End If
                    
                    If matchFound = True Then
                        Controls("Row1" & x).BackColor = vbYellow
                    Else
                        Controls("Row1" & x).BackColor = vbRed
                    End If
                End If
            Else
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row1" & x).BackColor = vbGreen
                Else
                    Controls("Row1" & x).BackColor = vbRed
                End If
            End If
        Next
    
    Case 2
        Row21.Caption = Cells(cardGuessedNo, 1)
        Row22.Caption = Cells(cardGuessedNo, 2)
        Row23.Caption = Cells(cardGuessedNo, 3)
        Row24.Caption = Cells(cardGuessedNo, 4)
        Row25.Caption = Cells(cardGuessedNo, 5)
        Row26.Caption = Cells(cardGuessedNo, 6)
        Row27.Caption = Cells(cardGuessedNo, 7)
        Row28.Caption = Cells(cardGuessedNo, 8)
        Row29.Caption = Cells(cardGuessedNo, 9)
        Row210.Caption = Cells(cardGuessedNo, 10)
        
        For x = 1 To 10
            matchFound = False
            If (x = 5 Or x = 9) Then
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row2" & x).BackColor = vbGreen
                Else
                    isTitle = InStr(1, Cells(cardGuessedNo, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATGuess = Cells(cardGuessedNo, x)
                       splitGuessLength = 0
                    Else
                        ATSplitGuess = Split(Cells(cardGuessedNo, x), ", ")
                        splitGuessLength = UBound(ATSplitGuess) - LBound(ATSplitGuess)
                    End If
                    
                    isTitle = InStr(1, Cells(cardChosen, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATChosen = Cells(cardChosen, x)
                       splitChosenlength = 0
                    Else
                        ATSplitChosen = Split(Cells(cardChosen, x), ", ")
                        splitChosenlength = UBound(ATSplitChosen) - LBound(ATSplitChosen)
                    End If
                    
                    If (splitGuessLength > 0 And splitChosenlength > 0) Then
                        For y = 0 To splitGuessLength
                            For Z = 0 To splitChosenlength
                                If ATSplitGuess(y) = ATSplitChosen(Z) Then
                                    matchFound = True
                                End If
                            Next
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength > 0) Then
                        For Z = 0 To splitChosenlength
                            If ATGuess = ATSplitChosen(Z) Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength > 0 And splitChosenlength = 0) Then
                        For y = 0 To splitGuessLength
                            If ATSplitGuess(y) = ATChosen Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength = 0) Then
                        If ATGuess = ATChosen Then
                            matchFound = True
                        End If
                    End If
                    
                    If matchFound = True Then
                        Controls("Row2" & x).BackColor = vbYellow
                    Else
                        Controls("Row2" & x).BackColor = vbRed
                    End If
                End If
            Else
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row2" & x).BackColor = vbGreen
                Else
                    Controls("Row2" & x).BackColor = vbRed
                End If
            End If
        Next
    
    Case 3
        Row31.Caption = Cells(cardGuessedNo, 1)
        Row32.Caption = Cells(cardGuessedNo, 2)
        Row33.Caption = Cells(cardGuessedNo, 3)
        Row34.Caption = Cells(cardGuessedNo, 4)
        Row35.Caption = Cells(cardGuessedNo, 5)
        Row36.Caption = Cells(cardGuessedNo, 6)
        Row37.Caption = Cells(cardGuessedNo, 7)
        Row38.Caption = Cells(cardGuessedNo, 8)
        Row39.Caption = Cells(cardGuessedNo, 9)
        Row310.Caption = Cells(cardGuessedNo, 10)
        
        For x = 1 To 10
            matchFound = False
            If (x = 5 Or x = 9) Then
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row3" & x).BackColor = vbGreen
                Else
                    isTitle = InStr(1, Cells(cardGuessedNo, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATGuess = Cells(cardGuessedNo, x)
                       splitGuessLength = 0
                    Else
                        ATSplitGuess = Split(Cells(cardGuessedNo, x), ", ")
                        splitGuessLength = UBound(ATSplitGuess) - LBound(ATSplitGuess)
                    End If
                    
                    isTitle = InStr(1, Cells(cardChosen, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATChosen = Cells(cardChosen, x)
                       splitChosenlength = 0
                    Else
                        ATSplitChosen = Split(Cells(cardChosen, x), ", ")
                        splitChosenlength = UBound(ATSplitChosen) - LBound(ATSplitChosen)
                    End If
                    
                    If (splitGuessLength > 0 And splitChosenlength > 0) Then
                        For y = 0 To splitGuessLength
                            For Z = 0 To splitChosenlength
                                If ATSplitGuess(y) = ATSplitChosen(Z) Then
                                    matchFound = True
                                End If
                            Next
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength > 0) Then
                        For Z = 0 To splitChosenlength
                            If ATGuess = ATSplitChosen(Z) Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength > 0 And splitChosenlength = 0) Then
                        For y = 0 To splitGuessLength
                            If ATSplitGuess(y) = ATChosen Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength = 0) Then
                        If ATGuess = ATChosen Then
                            matchFound = True
                        End If
                    End If
                    
                    If matchFound = True Then
                        Controls("Row3" & x).BackColor = vbYellow
                    Else
                        Controls("Row3" & x).BackColor = vbRed
                    End If
                End If
            Else
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row3" & x).BackColor = vbGreen
                Else
                    Controls("Row3" & x).BackColor = vbRed
                End If
            End If
        Next
    
    Case 4
        Row41.Caption = Cells(cardGuessedNo, 1)
        Row42.Caption = Cells(cardGuessedNo, 2)
        Row43.Caption = Cells(cardGuessedNo, 3)
        Row44.Caption = Cells(cardGuessedNo, 4)
        Row45.Caption = Cells(cardGuessedNo, 5)
        Row46.Caption = Cells(cardGuessedNo, 6)
        Row47.Caption = Cells(cardGuessedNo, 7)
        Row48.Caption = Cells(cardGuessedNo, 8)
        Row49.Caption = Cells(cardGuessedNo, 9)
        Row410.Caption = Cells(cardGuessedNo, 10)
        
        For x = 1 To 10
            matchFound = False
            If (x = 5 Or x = 9) Then
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row4" & x).BackColor = vbGreen
                Else
                    isTitle = InStr(1, Cells(cardGuessedNo, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATGuess = Cells(cardGuessedNo, x)
                       splitGuessLength = 0
                    Else
                        ATSplitGuess = Split(Cells(cardGuessedNo, x), ", ")
                        splitGuessLength = UBound(ATSplitGuess) - LBound(ATSplitGuess)
                    End If
                    
                    isTitle = InStr(1, Cells(cardChosen, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATChosen = Cells(cardChosen, x)
                       splitChosenlength = 0
                    Else
                        ATSplitChosen = Split(Cells(cardChosen, x), ", ")
                        splitChosenlength = UBound(ATSplitChosen) - LBound(ATSplitChosen)
                    End If
                    
                    If (splitGuessLength > 0 And splitChosenlength > 0) Then
                        For y = 0 To splitGuessLength
                            For Z = 0 To splitChosenlength
                                If ATSplitGuess(y) = ATSplitChosen(Z) Then
                                    matchFound = True
                                End If
                            Next
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength > 0) Then
                        For Z = 0 To splitChosenlength
                            If ATGuess = ATSplitChosen(Z) Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength > 0 And splitChosenlength = 0) Then
                        For y = 0 To splitGuessLength
                            If ATSplitGuess(y) = ATChosen Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength = 0) Then
                        If ATGuess = ATChosen Then
                            matchFound = True
                        End If
                    End If
                    
                    If matchFound = True Then
                        Controls("Row4" & x).BackColor = vbYellow
                    Else
                        Controls("Row4" & x).BackColor = vbRed
                    End If
                End If
            Else
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row4" & x).BackColor = vbGreen
                Else
                    Controls("Row4" & x).BackColor = vbRed
                End If
            End If
        Next
    
    Case 5
        Row51.Caption = Cells(cardGuessedNo, 1)
        Row52.Caption = Cells(cardGuessedNo, 2)
        Row53.Caption = Cells(cardGuessedNo, 3)
        Row54.Caption = Cells(cardGuessedNo, 4)
        Row55.Caption = Cells(cardGuessedNo, 5)
        Row56.Caption = Cells(cardGuessedNo, 6)
        Row57.Caption = Cells(cardGuessedNo, 7)
        Row58.Caption = Cells(cardGuessedNo, 8)
        Row59.Caption = Cells(cardGuessedNo, 9)
        Row510.Caption = Cells(cardGuessedNo, 10)
        
        For x = 1 To 10
            matchFound = False
            If (x = 5 Or x = 9) Then
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row5" & x).BackColor = vbGreen
                Else
                    isTitle = InStr(1, Cells(cardGuessedNo, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATGuess = Cells(cardGuessedNo, x)
                       splitGuessLength = 0
                    Else
                        ATSplitGuess = Split(Cells(cardGuessedNo, x), ", ")
                        splitGuessLength = UBound(ATSplitGuess) - LBound(ATSplitGuess)
                    End If
                    
                    isTitle = InStr(1, Cells(cardChosen, x), ",")
                    If (isTitle = 0) Then 'there is no title
                       ATChosen = Cells(cardChosen, x)
                       splitChosenlength = 0
                    Else
                        ATSplitChosen = Split(Cells(cardChosen, x), ", ")
                        splitChosenlength = UBound(ATSplitChosen) - LBound(ATSplitChosen)
                    End If
                    
                    If (splitGuessLength > 0 And splitChosenlength > 0) Then
                        For y = 0 To splitGuessLength
                            For Z = 0 To splitChosenlength
                                If ATSplitGuess(y) = ATSplitChosen(Z) Then
                                    matchFound = True
                                End If
                            Next
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength > 0) Then
                        For Z = 0 To splitChosenlength
                            If ATGuess = ATSplitChosen(Z) Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength > 0 And splitChosenlength = 0) Then
                        For y = 0 To splitGuessLength
                            If ATSplitGuess(y) = ATChosen Then
                                matchFound = True
                            End If
                        Next
                    ElseIf (splitGuessLength = 0 And splitChosenlength = 0) Then
                        If ATGuess = ATChosen Then
                            matchFound = True
                        End If
                    End If
                    
                    If matchFound = True Then
                        Controls("Row5" & x).BackColor = vbYellow
                    Else
                        Controls("Row5" & x).BackColor = vbRed
                    End If
                End If
            Else
                If (Cells(cardGuessedNo, x) = Cells(cardChosen, x)) Then
                    Controls("Row5" & x).BackColor = vbGreen
                Else
                    Controls("Row5" & x).BackColor = vbRed
                End If
            End If
        Next
    
    Case Else
        MsgBox ("You've had all your guesses")
        If (Cells(cardChosen, 2) <> "-") Then
            MsgBox ("The answer was " & Cells(cardChosen, 1) & ", " & Cells(cardChosen, 2))
        Else
            MsgBox ("The answer was " & Cells(cardChosen, 1))
        End If
    End Select
    
    GuessNo = GuessNo + 1
    
End Sub

Private Sub CommandButton2_Click()
    Randomize
    cardChosen = Int((252 * Rnd) + 1)
    cardChosen = cardChosen + 1
    GuessNo = 1
    
    For x = 1 To 5
        For y = 1 To 10
            Controls("Row" & x & y).BackColor = vbWhite
            Controls("Row" & x & y).Caption = ""
        Next
    Next
End Sub
