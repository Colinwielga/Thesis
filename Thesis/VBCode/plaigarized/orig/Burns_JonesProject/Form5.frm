VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14670
   LinkTopic       =   "Form5"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   14295
      Left            =   0
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   14235
      ScaleWidth      =   19035
      TabIndex        =   0
      Top             =   0
      Width           =   19095
      Begin VB.CommandButton cmdWomens 
         Caption         =   "How Will The Women's Champions Be Remembered?"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   14160
         TabIndex        =   7
         Top             =   9720
         Width           =   2895
      End
      Begin VB.CommandButton cmdHowManyWins 
         Caption         =   "How Will The Men's Champions Be Remembered?"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   10920
         TabIndex        =   6
         Top             =   11280
         Width           =   2895
      End
      Begin VB.PictureBox picWomensWinners 
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   6915
         TabIndex        =   5
         Top             =   2160
         Width           =   6975
      End
      Begin VB.PictureBox picMensWinners 
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   6915
         TabIndex        =   4
         Top             =   0
         Width           =   6975
      End
      Begin VB.CommandButton cmdWomensChampions 
         Caption         =   "The Women Champions Are....?!"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7920
         TabIndex        =   3
         Top             =   9720
         Width           =   2895
      End
      Begin VB.CommandButton cmdMensChampions 
         Caption         =   "The Mens Champions Are....?!"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5160
         TabIndex        =   2
         Top             =   11280
         Width           =   2895
      End
      Begin VB.CommandButton cmdReturnToForum 
         Caption         =   "Return To Main Page"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1920
         TabIndex        =   1
         Top             =   9720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHowManyWins_Click()
Dim LastName As String, I As Integer
LastName = InputBox("Enter A Last Name of an Above Champion")
C = 0
For I = 1 To Ctr
    If MLastName(I) = LastName Then
        C = C + 1
    End If
Next I
picMensWinners.Cls
Select Case C
    Case 0
        picMensWinners.Print "Sorry you were no champion, better luck next time"
    Case 1, 2
        picMensWinners.Print "This player has showed perseverance through many matches to take home the title of Champion"
    Case 3, 4
        picMensWinners.Print "This Player is really close to breaking the records and establishing your self as a true champion"
    Case 5
        picMensWinners.Print "You are a legend, and you will be remembered in the sport for ever!"
End Select

'this button uses select case statements to allow the user to search through the male players and determine what characteristic they each have'
End Sub

Private Sub cmdMensChampions_Click()
    Open App.Path & "\Frenchopenmens.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, MFirstName(Ctr), MLastName(Ctr), MYear(Ctr)
    Loop
Close #1
MsgBox "The Mens Champions Have Been Entered"
picMensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picMensWinners.Print "*********************************************************"
For I = 1 To Ctr
    picMensWinners.Print MLastName(I), MFirstName(I); Tab(35); MYear(I)
Next I
    

'this button helps to bring in information from a text file on notepad and print it off in the picture box'

    
End Sub

Private Sub cmdReturnToForum_Click()
    Form5.Hide
    Form2.Show
    'this button takes the user back to the main menu'
End Sub

Private Sub cmdWomens_Click()
Dim LastName As String, I As Integer
LastName = InputBox("Enter A Last Name of an Above Champion")
C = 0
For I = 1 To Ctr
    If WLastName(I) = LastName Then
        C = C + 1
    End If
Next I
picWomensWinners.Cls
Select Case C
    Case 0
        picWomensWinners.Print "Sorry you were no champion, better luck next time"
    Case 1, 2
        picWomensWinners.Print "This player has showed perseverance through many matches to take home the title of Champion"
    Case 3, 4
        picWomensWinners.Print "This Player is really close to breaking the records and establishing your self as a true champion"
    Case 5
        picWomensWinners.Print "You are a legend, and you will be remembered in the sport for ever!"
End Select

'this button uses select case statements to allow the user to search through the female players and determine what characteristic they each have'
End Sub

Private Sub cmdWomensChampions_Click()
    Open App.Path & "\Frenchopenwomens.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, WFirstName(Ctr), WLastName(Ctr), WYear(Ctr)
    Loop
Close #1
MsgBox "The Womens Champions Have Been Entered"
picWomensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picWomensWinners.Print "*******************************************************"
For I = 1 To Ctr
    picWomensWinners.Print WLastName(I), WFirstName(I); Tab(35); WYear(I)
Next I
'this button helps to bring in information from a text file on notepad and print it off in the picture box'

End Sub
