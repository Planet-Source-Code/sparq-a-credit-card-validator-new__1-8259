VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CC# Verify - "
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   2580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   308
      TabIndex        =   3
      Text            =   "49927398716"
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   180
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   180
      TabIndex        =   1
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check CC #"
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPARQ - jay@alphamedia.net       "
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1380
      Width           =   2445
   End
   Begin VB.Image imgDisc 
      Height          =   450
      Left            =   1620
      Top             =   780
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgMC 
      Height          =   465
      Left            =   1560
      Top             =   780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgAmex 
      Height          =   465
      Left            =   1620
      Top             =   780
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image imgVisa 
      Height          =   465
      Left            =   1620
      Top             =   780
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Disclaimer - This code is for educational purposes ONLY. I obtained this formula from a website some time ago."
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HOW IT ALL WORKS:
'1. Double the value of alternate digits of the primary account
'   account number beginning with the 2nd digit from the right.
'   (The 1st digit from the right is the check digit)
'
'2. Add the individual digits comprising the products obtained
'   in Step 1 to each of the uneffected digits in the origional
'   account number
'
'3. The total obtained in Step 2 must be a number that is ending
'   in ZERO (i.e.   TOTAL mod 10 = 0)
'
'Using the test card: 49927398716,
'
'    4  9  9  2  7  3  9  8  7  1  6
'      x2    x2    x2    x2    x2
'    -------------------------------
'    4 18  9  4  7  6  9 16  7  2  6
'
'    4+1+8+9+4+7+6+9+1+6+7+2+6 = 70
'    70 mod 10 = 0
'    Card is valid.
'
'
'
' INFO YOU MAY NEED!
'
' CARD NAME    Prefix    Length   Check Digit Algorithm
' MasterCard   51-55     16       Mod 10 = 0
' Visa         4         13,16    Mod 10 = 0
' AmEx         34,37     15       Mod 10 = 0
' Discover     6011      16       Mod 10 = 0
' Diner Club   300-305   14       Mod 10 = 0
'              36,38


Private Sub Command1_Click()
        Dim GoodCC As Boolean
        Dim CardNum, c1, c2, c3, c4, c5, CardType As String
        Dim CardTotal As Integer
        List1.Clear
        List2.Clear
        imgMC.Visible = False
        imgVisa.Visible = False
        imgDisc.Visible = False
        imgAmex.Visible = False
        '--------------DELETE SPACES
        Do While InStr(1, Text1, " ") <> 0
          Text1 = Left$(Text1, InStr(1, Text1, " ") - 1) & Mid$(Text1, InStr(1, Text1, " ") + 1)
        Loop
        '--------------DELETE DASHES
        Do While InStr(1, Text1, "-") <> 0
          Text1 = Left$(Text1, InStr(1, Text1, "-") - 1) & Mid$(Text1, InStr(1, Text1, "-") + 1)
        Loop
        CardNum = Text1
        Dim X!, i!
        X = 2
        For i = Len(CardNum) To 1 Step -1
        
         If X \ 2 = X / 2 Then
          List1.AddItem Mid$(CardNum, i, 1)
          X = 1
         Else
          List1.AddItem Val(Mid$(CardNum, i, 1)) * 2
          X = 2
         End If
        Next i
        
        For i = 0 To List1.ListCount - 1
         If Len(List1.List(i)) = 1 Then
          List2.AddItem List1.List(i)
         Else
          List2.AddItem Left$(List1.List(i), 1)
          List2.AddItem Right$(List1.List(i), 1)
         End If
        Next i
        
        For i = 0 To List2.ListCount - 1
         CardTotal = CardTotal + List2.List(i)
        Next i
        
        If CardTotal Mod 10 = 0 Then
         GoodCC = True
        Else
         GoodCC = False
         Text1.SetFocus
        End If
        
        
        If GoodCC = True Then
          c1 = Val(Left$(Text1, 1))
          c2 = Val(Left$(Text1, 2))
          c3 = Val(Left$(Text1, 3))
          c4 = Val(Left$(Text1, 4))
         
          If c2 >= 51 And c2 <= 55 Then
            CardType = "Master Card - VALID!"
            imgMC.Visible = True

            If Len(Text1) <> 16 Then
              CardType = "Invalid MasterCard Format"
            End If
         
          ElseIf c1 = 4 Then
            CardType = "Visa - VALID!"
            imgVisa.Visible = True

            If Len(Text1) <> 13 And Len(Text1) <> 16 Then
              CardType = "Invalid Visa Format"
            End If
        
          ElseIf c2 = 34 Or c2 = 37 Then
            CardType = "American Express - VALID!"
            imgAmex.Visible = True

            If Len(Text1) <> 15 Then
              CardType = "Invalid American Express Format"
            End If
        
          ElseIf c4 = 6011 Then
            CardType = "Discover - VALID!"
            imgDisc.Visible = True
            If Len(Text1) <> 16 Then
                CardType = "Invalid Discover Format"
            End If
        
          Else
            CardType = "Not Valid Credit Card"
          End If
        
        End If
        MsgBox CardType
End Sub

Private Sub Form_Load()
    imgVisa.Picture = LoadPicture(App.Path & "\visa.gif")
    imgDisc.Picture = LoadPicture(App.Path & "\discover.gif")
    imgMC.Picture = LoadPicture(App.Path & "\mcard.gif")
    imgAmex.Picture = LoadPicture(App.Path & "\amex.gif")
End Sub

Private Sub Timer1_Timer()
            
            String1 = Right$(Label1.Caption, 1)
            String2 = Left$(Label1.Caption, Len(Label1.Caption) - 1)
            Label1 = String1 & String2
        
End Sub
