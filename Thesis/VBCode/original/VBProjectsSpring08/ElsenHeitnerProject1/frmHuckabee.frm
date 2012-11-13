VERSION 5.00
Begin VB.Form frmHuckabee 
   Caption         =   "Huckabee's Bio"
   ClientHeight    =   6990
   ClientLeft      =   3945
   ClientTop       =   2940
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   Picture         =   "frmHuckabee.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   9495
   Begin VB.PictureBox picMike 
      Height          =   1935
      Left            =   240
      Picture         =   "frmHuckabee.frx":3186
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Cantidates"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdAchievements 
      Caption         =   "Mike's Achievements"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Achievements by Year "
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "See Biography"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   5775
      Left            =   2400
      ScaleHeight     =   5715
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   960
      Width           =   6735
   End
End
Attribute VB_Name = "frmHuckabee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Huckabee's Bio (frmHuckabee.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 12, 2008
'PURPOSE:  This form is to show Mike Huckabee's Biography and achievements.

Option Explicit
Dim Ctr As Integer
Dim HuckabeeBio(1 To 100) As String
Dim Achievements(1 To 40) As String
Dim Year(1 To 40) As Single

'Displays Mike's Achievements and year of achievement by loading them from a text file
Private Sub cmdAchievements_Click()
cmdSort.Visible = True
Dim J As Integer

picResults.Cls

Open App.Path & "\BioTexts\HuckabeeAchievements.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Achievements(Ctr), Year(Ctr)
Loop
    
    picResults.Print "Achievement"; Tab(60); "Year"
    picResults.Print "*************************************************************************************"
For J = 1 To Ctr
    picResults.Print Achievements(J); Tab(60); Year(J)
Next J

Close #1

End Sub

'Takes you back to the Cantidates form
Private Sub cmdBack_Click()
frmHuckabee.Hide
frmCantidates.Show
End Sub

'Displays Mike's Biography by loading it from a text file
Private Sub cmdBio_Click()

Open App.Path & "\BioTexts\HuckabeeBio.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, HuckabeeBio(Ctr)
    picResults.Print HuckabeeBio(Ctr)
    
Loop
Close #1

End Sub

'Sorts Mike's Achievements by Year by using a bubble sort
Private Sub cmdSort_Click()
Dim Pos As Integer
Dim pass As Integer
Dim TempAchievements As String
Dim TempYear As Integer
Dim K As Single
picResults.Cls

For pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - pass
        If Year(Pos) > Year(Pos + 1) Then
            TempYear = Year(Pos)
            Year(Pos) = Year(Pos + 1)
            Year(Pos + 1) = TempYear
            
            TempAchievements = Achievements(Pos)
            Achievements(Pos) = Achievements(Pos + 1)
            Achievements(Pos + 1) = TempAchievements
        End If
    Next Pos
Next pass

picResults.Print "Year", "Achievements"
picResults.Print "*************************************************************************************"

For K = 1 To Ctr
    picResults.Print Year(K), Achievements(K)
Next K
End Sub
