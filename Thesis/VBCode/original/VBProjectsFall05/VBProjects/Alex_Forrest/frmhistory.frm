VERSION 5.00
Begin VB.Form frmhistory 
   BackColor       =   &H000040C0&
   Caption         =   "History of Rugby"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "search for part of history"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmddisplayentire 
      Caption         =   "display entire history"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox picbox 
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton cmdreturntomainmenui 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Baskerville MT for Brill 01 SC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   0
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Image imgRugby4 
      Height          =   2010
      Left            =   7800
      Picture         =   "frmhistory.frx":0000
      Top             =   2280
      Width           =   2595
   End
End
Attribute VB_Name = "frmhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : RugbyVBProject (Rugby.vbp)
'Form Name : frmhistory(frmhistory.frm)
'Author: Alex Forrest
'purpose of the form: This form is designed to give brief explanations of the history
    'of the game of rugby.  One button displays the entire history and another button is
    'designed to let the user input which era of rugby they wish to learn about through an
    'input box.

Private Sub cmddisplayentire_Click()
    picbox.Cls 'clears the picture box so the following information
    picbox.Print "The Beginning of Rugby" 'the following print commands print the beginning of rugby information
    picbox.Print "******************************************************************************************************"
    picbox.Print "  In 1800's formalities were introduced to football rules in the seven major public schools of England."
    picbox.Print "Six of the seven schools were largely playing the same game (including Eton, Harrow and Winchester),"
    picbox.Print "while the seventh Rugby School (founded in 1567) at Warwickshire, was playing a markedly different"
    picbox.Print "version of football."
    picbox.Print "  The Rugby Football Union's (RFU) much revered tale of how in 1823 the young Rugby School student,"
    picbox.Print "William Webb Ellis, in a fine disregard for the rules, picked up the ball and ran with it in a"
    picbox.Print "defining moment in sports history is now accepted by sports historians as the beginning of the"
    picbox.Print "sport of Rugby."
    picbox.Print
    picbox.Print "The Great Divide" 'the following print commands print the great divide of rugby
    picbox.Print "******************************************************************************************************"
    picbox.Print "  In 1895 the movement for the creation of a Northern Rugby Union outside of the control of the RFU"
    picbox.Print "had reached a crescendo. In one final effort to reign in the rising upheavel the RFU broadened its"
    picbox.Print "definition of professionalism to include playing on a ground where gate money was taken and/or any"
    picbox.Print "game to be played with less than 15 men-a-side. The RFU knew that some of the northern clubs had"
    picbox.Print "been contemplating reducing the number of players in teams to less than 15 to improve the crowd"
    picbox.Print "appeal - in fact the RFU had even considered the option itself in 1892."
    picbox.Print
    picbox.Print "  Thus the Great Divide of 1895 produced two new sports from the shared rugby parent, not the minor"
    picbox.Print "loss of an unimportant appendage as the RFU has forever since portrayed it.The split would also"
    picbox.Print "ensure that RU would forever polarise itself as a middle-class game and live its amatuer lie for"
    picbox.Print "a further hundred years. On 29 August 1895 twenty-one clubs met at the George Hotel in Huddersfield"
    picbox.Print "and formed the Northern Rugby Union (later to become known as Rugby League)."
    picbox.Print
    picbox.Print "Modern Rugby Era" 'the following print commands print the modern era of rugby
    picbox.Print "******************************************************************************************************"
    picbox.Print "  As Rugby has evolved throughout the years, it still continues to be played at the professional level"
    picbox.Print "today. Rugby is extremely popular on the european continent, but is still played worldwide. Rugby"
    picbox.Print "powerhouses include England, Australia, and New Zealand.  These teams, along with many others"
    picbox.Print "around the world, compete in the Rugby World Cup held every for years.  Aside from this, there are"
    picbox.Print "leagues all around the world that hold league matches and tournaments.  Rugby has, and continues to"
    picbox.Print "become a popular sport around the world."
End Sub

Private Sub cmdquit_Click()
    End 'quits the program
End Sub

Private Sub cmdreturntomainmenui_Click()
    frmhistory.Hide
    frmMainmenu.Show 'takes the user back to the main menu
End Sub

Private Sub cmdsearch_Click()
    Dim history As Integer
    history = InputBox("Input the era of Rugby you wish to learn about - Beginning = 1, Great Divide = 2, or Modern = 3") 'sets the history to the information inputted in the input box
    picbox.Cls 'clears the picture box
        If history = 1 Then 'tests the user's input in attempt to find a match
            picbox.Print "The Beginning of Rugby" 'prints the following information if the users inputs a 1
            picbox.Print "******************************************************************************************************"
            picbox.Print "  In 1800's formalities were introduced to football rules in the seven major public schools of England."
            picbox.Print "Six of the seven schools were largely playing the same game (including Eton, Harrow and Winchester),"
            picbox.Print "while the seventh Rugby School (founded in 1567) at Warwickshire, was playing a markedly different"
            picbox.Print "version of football."
            picbox.Print "  The Rugby Football Union's (RFU) much revered tale of how in 1823 the young Rugby School student,"
            picbox.Print "William Webb Ellis, in a fine disregard for the rules, picked up the ball and ran with it in a"
            picbox.Print "defining moment in sports history is now accepted by sports historians as the beginning of the"
            picbox.Print "sport of Rugby."
        ElseIf history = 2 Then 'tests the user's input in attempt to find a match
            picbox.Print "The Great Divide" 'prints the following information if the user inputs a 2
            picbox.Print "******************************************************************************************************"
            picbox.Print "  In 1895 the movement for the creation of a Northern Rugby Union outside of the control of the RFU"
            picbox.Print "had reached a crescendo. In one final effort to reign in the rising upheavel the RFU broadened its"
            picbox.Print "definition of professionalism to include playing on a ground where gate money was taken and/or any"
            picbox.Print "game to be played with less than 15 men-a-side. The RFU knew that some of the northern clubs had"
            picbox.Print "been contemplating reducing the number of players in teams to less than 15 to improve the crowd"
            picbox.Print "appeal - in fact the RFU had even considered the option itself in 1892."
            picbox.Print
            picbox.Print "  Thus the Great Divide of 1895 produced two new sports from the shared rugby parent, not the minor"
            picbox.Print "loss of an unimportant appendage as the RFU has forever since portrayed it.The split would also"
            picbox.Print "ensure that RU would forever polarise itself as a middle-class game and live its amatuer lie for"
            picbox.Print "a further hundred years. On 29 August 1895 twenty-one clubs met at the George Hotel in Huddersfield"
            picbox.Print "and formed the Northern Rugby Union (later to become known as Rugby League)."
        ElseIf history = 3 Then 'tests the user's input in attempt to find a match
            picbox.Print "Modern Rugby Era" 'prints the following information if the user inputs a 3
            picbox.Print "******************************************************************************************************"
            picbox.Print "  As Rugby has evolved throughout the years, it still continues to be played at the professional level"
            picbox.Print "today. Rugby is extremely popular on the european continent, but is still played worldwide. Rugby"
            picbox.Print "powerhouses include England, Australia, and New Zealand.  These teams, along with many others"
            picbox.Print "around the world, compete in the Rugby World Cup held every for years.  Aside from this, there are"
            picbox.Print "leagues all around the world that hold league matches and tournaments.  Rugby has, and continues to"
            picbox.Print "become a popular sport around the world."
        Else 'instructs the program to print the folloing information in the case that a match of the user's input was not found
            MsgBox "Sorry, You did not input the correct data!", , "Error"
        End If 'ends the if statement
End Sub
