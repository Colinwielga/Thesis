VERSION 5.00
Begin VB.Form frmBookCase 
   BackColor       =   &H80000007&
   Caption         =   "Book Case"
   ClientHeight    =   8610
   ClientLeft      =   1350
   ClientTop       =   1725
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   9990
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   7680
      Picture         =   "frmBookCase.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.PictureBox picBook3 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   7680
      Picture         =   "frmBookCase.frx":9024
      ScaleHeight     =   2175
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   5640
      Width           =   1455
   End
   Begin VB.PictureBox picBookOne 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   7680
      Picture         =   "frmBookCase.frx":1211C
      ScaleHeight     =   2175
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   600
      Picture         =   "frmBookCase.frx":12A5D
      ScaleHeight     =   4455
      ScaleWidth      =   5655
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdOpenBook 
      Caption         =   "Read It?"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdKeepReading 
      Caption         =   "Continue Reading"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Finished Looking"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.OptionButton optBookFive 
      BackColor       =   &H8000000C&
      Caption         =   "Book Three"
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   8040
      Width           =   1335
   End
   Begin VB.OptionButton optBookThree 
      BackColor       =   &H8000000C&
      Caption         =   "Book Two"
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.OptionButton optBookOne 
      BackColor       =   &H8000000C&
      Caption         =   "Book One"
      Height          =   255
      Left            =   7800
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox picBookCasetxt 
      Height          =   3015
      Left            =   480
      ScaleHeight     =   2955
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   4800
      Width           =   6135
   End
End
Attribute VB_Name = "frmBookCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim diarystring(1 To 4) As String
Dim diarypos As Integer

Private Sub cmdExit_Click()
Dim answer As Integer

answer = MsgBox("Are you sure you are done looking at the bookcase?", vbYesNo)
    If answer = vbYes Then
        frmLibrary.Show
        frmBookCase.Hide
    End If
End Sub

Private Sub cmdKeepReading_Click()
    '************************
    Dim Color(1 To 8) As String
    Dim Letter(1 To 8) As String           'Book 3 Variables
    Dim ColorCTR As Integer
    Dim ColorPOS As Integer
    '************************
    If optBookOne = True Then
    diarypos = diarypos + 1
        If diarypos < 5 Then
             Dim FileString As String, stringLength As Integer, tempString As String, linelength As Integer, i As Integer
             Dim space As String   'Variables from Chris Kerber's Code
             Dim pos As Integer
             picBookCasetxt.Cls
             
             diarystring(1) = "June 13th, ’86 Those military fools are watching over my back again. They want to use the GMV for warfare. To kill people! I am just trying to save people. I am disgusted with these people- so impatient. If I finish the GMV, I can save my love. That’s all that matters..."
             diarystring(2) = "June 30th, ’86  There was another military bombardment on the shelter, few leaking pipes but everything still intact. My dear’s symptoms are getting worse by the hour- I placed my love into a medically-induced coma to take end the pain. I said goodbye but not for the last time."
             diarystring(3) = "August 15th , ’86  The GMV is genius! This can end all of man’s problems! My virus is programmed to invade the genetic sectors of cells and change the genetic code. I can change the function of cells or destroy them. I hope that it will be able to be program it to find and kill cancer cells and restore damaged brain cells and nerves. Soon..."
             diarystring(4) = "September 1st, ’86 Today is the day. The first of September is a day to remember. I will test the GMV on my love and destroy the cancer.  Mankind will be able to overcome any biological hurdle capable! I am the pioneer of Humankind’s  future!"
             
             linelength = 80
             
             
            
             space = " "
             pos = 0
             
                 
                   FileString = diarystring(diarypos)
                     stringLength = Len(FileString)                'length of that long string
                     i = 1                                       'basically a counter
                     linelength = 80
                     While i + linelength < stringLength          'write another line until counter > stringLength
                        tempString = Mid(FileString, i, linelength) 'computes the next line to write using the mid function
                        pos = InStrRev(tempString, space)
                        tempString = Mid(FileString, i, pos)
                        picBookCasetxt.Print tempString              'prints that computed line
                        i = i + pos                       'increments the counter by the lineLenth
                     Wend
                     picBookCasetxt.Print Right(FileString, stringLength - i + 1) 'prints the last line which is left out in the loop.
         Else
         cmdKeepReading.Enabled = False
         End If
        
    End If
    



    If optBookThree = True Then
        picBookCasetxt.Cls
        Open App.Path & "\colorcodex.txt" For Input As #1
        ColorCTR = 0
        Do Until EOF(1)         'Saving to an Array
            ColorCTR = ColorCTR + 1
            Input #1, Color(ColorCTR), Letter(ColorCTR)
        Loop
        Close #1
        ColorPOS = 0
        For ColorPOS = 1 To ColorCTR
            picBookCasetxt.Print Color(ColorPOS), Letter(ColorPOS)
        Next ColorPOS
    
        picBookCasetxt.Print " "
        picBookCasetxt.Print " "
        picBookCasetxt.Print " "
        picBookCasetxt.Print " "
        picBookCasetxt.Print " This information looks like it could be helpful. Maybe you should write it down..."
        cmdKeepReading.Enabled = False
        
    End If
End Sub

Private Sub cmdOpenBook_Click()
    cmdOpenBook.Enabled = False
    If optBookFive = True Then
    cmdKeepReading.Enabled = False
        If EmblemTwo = True Then
            picBookCasetxt.Print "You have already taken the Emblem Piece from the book."
            picBookCasetxt.Print "There is nothing left."
        Else
            picBookCasetxt.Print " "
            picBookCasetxt.Print " "
            picBookCasetxt.Print "You opened the book..."
            picBookCasetxt.Print " "
            picBookCasetxt.Print "To your surprise the book is empty! It is a hollowed out book holding"
            picBookCasetxt.Print "a Piece of the BROKEN EMBLEM!"
            picBookCasetxt.Print "You hold on to the Emblem Piece for future use. Only ONE more piece of the "
            picBookCasetxt.Print "Emblem remains."
            EmblemTwo = True
        End If
        
    End If
    '*********************************
    If optBookThree = True Then
        picBookCasetxt.Cls
        picBookCasetxt.Print "E-Mail from Ryu Moriyama"
        picBookCasetxt.Print "MSG TITLE: Color-Codex Change (Don't Delete!)"
        picBookCasetxt.Print "MSG:"
        picBookCasetxt.Print " "
        picBookCasetxt.Print "Hey Everybody"
        picBookCasetxt.Print " "
        picBookCasetxt.Print "The color codex system is being reset for a new match. As you know, each color has"
        picBookCasetxt.Print "a particular Alphabetical Equivalent to be used. This change is to be effective "
        picBookCasetxt.Print "immediately."
        picBookCasetxt.Print "The list is attached."
        picBookCasetxt.Print " "
        picBookCasetxt.Print "Please memorize this information."
        picBookCasetxt.Print " "
        picBookCasetxt.Print "Thanks"
        picBookCasetxt.Print "Ryu Moriyama"
        
        cmdKeepReading.Enabled = True 'Allows user to see attachment
        
    End If
    '***********************************
    
    Dim openstring As String
    Dim pos As Integer
    
    If optBookOne = True Then
        picBookCasetxt.Cls
        'Open App.Path & "\diary.txt" For Input As #2
        'diaryctr = 0
        'Do While Not EOF(2)
        '    diaryctr = diaryctr + 1
        '    Input #2, diarystring(diaryctr)
        'Loop
        'Close #2
        
    Dim FileString As String, stringLength As Integer, tempString As String, linelength As Integer, i As Integer
    Dim space As String   'Variables from Chris Kerber's Code
    space = " "
    pos = 0
    linelength = 80
    
    openstring = "June 6th,  ’86 – What a great day it has been today. I am happy to say that the first prototype of the genetic-mutator virus (GMV) has been completed. Soon I will be able to program complex codes in it! Maybe I will be able to cure my love’s cancer before it is too late."
    stringLength = Len(openstring)                'length of that long ass string
    i = 1                                       'basically a counter
    While i + linelength < stringLength          'write another line until counter > stringLength
       tempString = Mid(openstring, i, linelength) 'computes the next line to write using the mid function
       pos = InStrRev(tempString, space)
       tempString = Mid(openstring, i, pos)
       picBookCasetxt.Print tempString              'prints that computed line
       i = i + pos                       'increments the counter by the lineLenth
    Wend
    picBookCasetxt.Print Right(openstring, stringLength - i + 1) 'prints the last line which is left out in the loop.
    diarypos = 0
    cmdKeepReading.Enabled = True
    
    End If
        
    
    
End Sub

Private Sub Form_activate()
    picBookCasetxt.Cls
    picBookCasetxt.Print "The bookcase is filled with many books, some are in different languages"
    picBookCasetxt.Print "Most of the books have worn edges from use."
    
    'certain buttons are not enabled/visible
    cmdOpenBook.Enabled = False
    cmdKeepReading.Enabled = False

End Sub

Private Sub optBookFive_Click()
    picBookCasetxt.Cls
    If EmblemTwo = True Then
        picBookCasetxt.Print "This is the hollowed out book. It is empty."
        
    Else
        picBookCasetxt.Print "The cover has been left blank and it is in perfect condition unlike the others"
        picBookCasetxt.Print "It even weighs much more than a normal book should. How odd..."
        
    End If
    cmdOpenBook.Enabled = True
    

End Sub

Private Sub optBookFour_Click()
picBookCasetxt.Cls
cmdOpenBook.Enabled = True
End Sub

Private Sub optBookOne_Click()
    picBookCasetxt.Cls
    picBookCasetxt.Print "You find a diary in the bookcase. You see that name '" & FirstName & "' written on"
    picBookCasetxt.Print "the inside cover. It is all handwritten and looks very personal."
    
    cmdOpenBook.Enabled = True
End Sub

Private Sub optBookThree_Click()
    picBookCasetxt.Cls
    picBookCasetxt.Print "This binder has the Japanese symbol for 'color' on it."
    picBookCasetxt.Print "You pause for a moment...         How did you know that?"
    picBookCasetxt.Print "What other secrets are hidden in the haze that covers your mind's eye?"
    picBookCasetxt.Print " "
    picBookCasetxt.Print " "
    picBookCasetxt.Print "The binder contains several email print outs. The last being"
    picBookCasetxt.Print "the most up-to-date."
    
    cmdOpenBook.Enabled = True
End Sub

Private Sub optBookTwo_Click()
picBookCasetxt.Cls
End Sub
