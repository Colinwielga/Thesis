VERSION 5.00
Begin VB.Form Form6
   BackColor       =   &H00000000&
   Caption         =   "Form6"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form6"
   ScaleHeight     =   10170
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1
      Height          =   10575
      Left            =   -240
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   10515
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   480
      Width           =   14055
      Begin VB.PictureBox picFast
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3240
         ScaleHeight     =   1755
         ScaleWidth      =   2235
         TabIndex        =   8
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton cmdHowFast
         Caption         =   $"Form6.frx":15F81
         BeginProperty Font
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   10440
         TabIndex        =   7
         Top             =   7080
         Width           =   2775
      End
      Begin VB.TextBox txtFast
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         TabIndex        =   6
         Top             =   4560
         Width           =   2055
      End
      Begin VB.PictureBox picWomensWinners
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         ScaleHeight     =   1875
         ScaleWidth      =   5595
         TabIndex        =   5
         Top             =   2280
         Width           =   5655
      End
      Begin VB.PictureBox picMensWinners
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   360
         ScaleHeight     =   1995
         ScaleWidth      =   5595
         TabIndex        =   4
         Top             =   120
         Width           =   5655
      End
      Begin VB.CommandButton cmdWomensChampions
         Caption         =   "The Women Champions Are...?!"
         BeginProperty Font
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   10440
         TabIndex        =   3
         Top             =   4920
         Width           =   2775
      End
      Begin VB.CommandButton cmdMensChampions
         Caption         =   "The Mens Champions Are...?!"
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
         Left            =   10440
         TabIndex        =   2
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CommandButton cmdTakeMeToForm2
         Caption         =   "Take me back to the main page."
         BeginProperty Font
            Name            =   "Footlight MT Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   10440
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdHowFast_Click()
' stuff
    Dim Fast As Integer, Equation As Integer
    Fast = txtFast.Text
    Equation = 1.609344 * Fast
    ' blah
    picFast.Print Equation, "KPH"

    'this button uses a numeric function to convert mph to kilometers per hour'

End Sub

Private Sub cmdMensChampions_Click()
        Open App.Path & "\Australianopenmens.txt" For Input As #1
        Ctr = o
        Do While Not EOF(1)
            Ctr = 1 + Ctr
            Input #1, MFirstName(Ctr), MLastName(Ctr), MYear(Ctr)
        Loop
    Close #1
    ' heylasdk fsd kf sdlfk sdf
    MsgBox "The Mens Champions Have Been Entered!"
    picMensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
    picMensWinners.Print "*********************************************************"
    For I = 1 To Ctr
        picMensWinners.Print MLastName(I), MFirstName(I); Tab(35); MYear(I)
    Next I



    'this button helps to bring in information from a text file on notepad and print it off in the picture box'



End Sub

Private Sub cmdTakeMeToForm2_Click()
    Form6.Hide
    Form2.Show
  'this takes the user back to the main menu'
End Sub

Private Sub cmdWomensChampions_Click()
        Open App.Path & "\Australianopenwomens.txt" For Input As #1
        Ctr = o
        Do While Not EOF(1)
            Ctr = 1 + Ctr
            ' alsdkfj sdlkf jdslfk sdj
            Input #1, WFirstName(Ctr), WLastName(Ctr), WYear(Ctr)
        Loop
    MsgBox "The Womens Champions Have Been Entered!"
    picWomensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
    picWomensWinners.Print "*******************************************************"
    Close #1
    For I = 1 To Ctr
        picWomensWinners.Print WLastName(I), WFirstName(I); Tab(35); WYear(I)
    Next I
    'this button helps to bring in information from a text file on notepad and print it off in the picture box'
End Sub

