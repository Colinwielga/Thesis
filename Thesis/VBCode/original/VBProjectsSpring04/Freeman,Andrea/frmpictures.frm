VERSION 5.00
Begin VB.Form AndreaFreemanfrmpictures 
   BackColor       =   &H00C0C000&
   Caption         =   "Pictures"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults5 
      Height          =   2655
      Left            =   7560
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox picResults4 
      Height          =   2655
      Left            =   3840
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   5640
      Width           =   2415
   End
   Begin VB.PictureBox picResults3 
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox picResults2 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Analysis "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      Picture         =   "frmpictures.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdhair 
      Caption         =   "See your    Hairstyle                       "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7920
      Picture         =   "frmpictures.frx":12CD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdmood 
      Caption         =   "See your Current Mood"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Picture         =   "frmpictures.frx":38EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmddream 
      Caption         =   "See your Dream Vacation"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      Picture         =   "frmpictures.frx":5F0B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdcolor 
      Caption         =   "See your Favorite Color"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Picture         =   "frmpictures.frx":852A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdanimal 
      Caption         =   "See your Favorite Animal"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Picture         =   "frmpictures.frx":AB49
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "AndreaFreemanfrmpictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmpictures (frmpictures.frm)
'Author: Andrea Freeman
'Date Written: March 15, 2004
'Purpose of Form: This form displays all of the pictures that the user chose and then
                  'returns them to the previous form, Analysis.


Dim Names(1 To 5) As String
Dim Pictures(1 To 5) As String

Private Sub cmdanimal_Click()

Dim Z As Integer

PATH = "N:\CS130\handin\Freeman, Andrea\"

Open PATH & "Animals.txt" For Input As #1

For Z = 1 To 5
    Input #1, Names(Z), Pictures(Z)
    If FavoriteAnimal(I) = Names(Z) Then
        picResults.Picture = LoadPicture("M:\CS130\Freeman, Andrea\" & Pictures(Z))
    End If
Next Z

Close #1
End Sub


Private Sub cmdcolor_Click()

Dim B As Integer

PATH = "N:\CS130\handin\Freeman, Andrea\"

Open PATH & "Colors.txt" For Input As #1

For B = 1 To 5
    Input #1, Names(B), Pictures(B)
    If FavoriteColor(J) = Names(B) Then
        picResults2.Picture = LoadPicture("M:\CS130\Freeman, Andrea\" & Pictures(B))
    End If
Next B

Close #1
End Sub

Private Sub cmddream_Click()

Dim C As Integer

PATH = "N:\CS130\handin\Freeman, Andrea\"

Open PATH & "Vacations.txt" For Input As #1

For C = 1 To 5
    Input #1, Names(C), Pictures(C)
    If DreamVacation(K) = Names(C) Then
        picResults3.Picture = LoadPicture("M:\CS130\Freeman, Andrea\" & Pictures(C))
    End If
Next C

Close #1
End Sub

Private Sub cmdhair_Click()

Dim F As Integer

PATH = "N:\CS130\handin\Freeman, Andrea\"

Open PATH & "Hairstyles.txt" For Input As #1

For F = 1 To 5
    Input #1, Names(F), Pictures(F)
    If Hairstyle(M) = Names(F) Then
        picResults5.Picture = LoadPicture("M:\CS130\Freeman, Andrea\" & Pictures(F))
    End If
Next F

Close #1
End Sub

Private Sub cmdmood_Click()

Dim E As Integer

PATH = "N:\CS130\handin\Freeman, Andrea\"

Open PATH & "Moods.txt" For Input As #1

For E = 1 To 5
    Input #1, Names(E), Pictures(E)
    If Mood(L) = Names(E) Then
        picResults4.Picture = LoadPicture("M:\CS130\Freeman, Andrea\" & Pictures(E))
    End If
Next E

Close #1
End Sub

Private Sub cmdreturn_Click()
'Hide the pictures form and return to the Analysis form.
AndreaFreemanfrmpictures.Hide
AndreaFreemanfrmAnalysis.Show

'Clear the picture boxes for repeated use.
picResults.Cls
picResults2.Cls
picResults3.Cls
picResults4.Cls
picResults5.Cls
End Sub

Private Sub Form_Load()

End Sub
