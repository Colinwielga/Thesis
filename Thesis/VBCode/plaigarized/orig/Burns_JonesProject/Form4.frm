VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   12855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form4"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortWomen 
      Caption         =   "Sort the female competitors by last name."
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
      Left            =   13080
      TabIndex        =   7
      Top             =   9600
      Width           =   3255
   End
   Begin VB.CommandButton cmdSortByLastName 
      Caption         =   "Sort the male competitors by last name."
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
      Left            =   13080
      TabIndex        =   6
      Top             =   7440
      Width           =   3255
   End
   Begin VB.CommandButton cmdWomensTitles 
      Caption         =   "I wonder who won the women's titles?"
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
      Left            =   13080
      TabIndex        =   3
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton cmdMensWinners 
      Caption         =   "I wonder who won the men's titles?"
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
      Left            =   13080
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturnToForum 
      Caption         =   "Return To Main Page!"
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
      Left            =   13080
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   11535
      Left            =   240
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   11475
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   480
      Width           =   12015
      Begin VB.PictureBox picWomensWinners 
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   5595
         TabIndex        =   5
         Top             =   2520
         Width           =   5655
      End
      Begin VB.PictureBox picMensWinners 
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2475
         ScaleWidth      =   5595
         TabIndex        =   4
         Top             =   0
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMensWinners_Click()
    Open App.Path & "\Wimbledonmens.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, MFirstName(Ctr), MLastName(Ctr), MYear(Ctr)
    Loop
Close #1
MsgBox "The Mens Champions Have Been Entered!"
picMensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picMensWinners.Print "*********************************************************"
For I = 1 To Ctr
    picMensWinners.Print MLastName(I), MFirstName(I); Tab(35); MYear(I)
Next I
    
'this button helps the user to print off the files from a text file on notepad'


    
End Sub

Private Sub cmdReturnToForum_Click()
    Form4.Hide
    Form2.Show
    'this takes the user back to the main menu'
End Sub

Private Sub cmdSortByLastName_Click()
Dim Pass As Integer, Pos As Integer, Temp As String, I As Integer
For Pass = 1 To Ctr
    For Pos = 1 To Ctr - 1
        If MLastName(Pos) > MLastName(Pos + 1) Then
            Temp = MLastName(Pos)
            MLastName(Pos) = MLastName(Pos + 1)
            MLastName(Pos + 1) = Temp
            
            Temp = MFirstName(Pos)
            MFirstName(Pos) = MFirstName(Pos + 1)
            MFirstName(Pos + 1) = Temp
            
            Temp = MYear(Pos)
            MYear(Pos) = MYear(Pos + 1)
            MYear(Pos + 1) = Temp
        End If
    Next Pos
Next Pass
picMensWinners.Cls
picMensWinners.Print "Last Name", "First Name", "Year"
picMensWinners.Print "***********************************************"
For I = 1 To Ctr
    picMensWinners.Print MLastName(I), MFirstName(I), MYear(I)
Next I
   'this button helps to sort all of the men's champions names by last name'
End Sub

Private Sub cmdSortWomen_Click()
Dim Pass As Integer, Pos As Integer, Temp As String, I As Integer
For Pass = 1 To Ctr
    For Pos = 1 To Ctr - 1
        If WLastName(Pos) > WLastName(Pos + 1) Then
            Temp = WLastName(Pos)
            WLastName(Pos) = WLastName(Pos + 1)
            WLastName(Pos + 1) = Temp
            
            Temp = WFirstName(Pos)
            WFirstName(Pos) = WFirstName(Pos + 1)
            WFirstName(Pos + 1) = Temp
            
            Temp = WYear(Pos)
            WYear(Pos) = WYear(Pos + 1)
            WYear(Pos + 1) = Temp
        End If
    Next Pos
Next Pass
picWomensWinners.Cls
picWomensWinners.Print "Last Name", "First Name", "Year"
picWomensWinners.Print "***********************************************"
For I = 1 To Ctr
    picWomensWinners.Print WLastName(I), WFirstName(I), WYear(I)
Next I
  'this button helps to sort the women's champions by last name as well'
End Sub

Private Sub cmdWomensTitles_Click()
    Open App.Path & "\Wimbledonwomens.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, WFirstName(Ctr), WLastName(Ctr), WYear(Ctr)
    Loop
Close #1
MsgBox "The Womens Champions Have Been Entered!"
picWomensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picWomensWinners.Print "*******************************************************"
For I = 1 To Ctr
    picWomensWinners.Print WLastName(I), WFirstName(I); Tab(35); WYear(I)
Next I
    
'this button helps to bring in information from a text file on notepad and print it off in the picture box'
    
    
End Sub
