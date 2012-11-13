VERSION 5.00
Begin VB.Form frmObama 
   Caption         =   "Obama's Bio"
   ClientHeight    =   6930
   ClientLeft      =   3750
   ClientTop       =   2940
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   Picture         =   "frmObama.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   9435
   Begin VB.PictureBox PicBarack 
      Height          =   1935
      Left            =   240
      Picture         =   "frmObama.frx":3186
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Achievements by Year"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAchievements 
      Caption         =   "Barack's Achievements"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "See Biography"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   5535
      Left            =   2400
      ScaleHeight     =   5475
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Cantidates"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "frmObama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Obama's Bio (frmObama.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 12, 2008
'PURPOSE:  This form is to show Barack Obama's Biography and achievements.

Option Explicit
Dim ObamaBio(1 To 100) As String
Dim Ctr As Integer
Dim Achievements(1 To 40) As String
Dim Year(1 To 40) As Single

'Displays Barack's Achievements and year of achievement by loading them from a text file
Private Sub cmdAchievements_Click()
cmdSort.Visible = True
Dim J As Integer

picResults.Cls

Open App.Path & "\BioTexts\BarackAchievements.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Achievements(Ctr), Year(Ctr)
Loop
    
    picResults.Print "Achievement"; Tab(80); "Year"
    picResults.Print "*******************************************************************************************************************"
For J = 1 To Ctr
    picResults.Print Achievements(J); Tab(80); Year(J)
Next J

Close #1

End Sub

'Takes you back to the Cantidates form
Private Sub cmdBack_Click()
frmObama.Hide
frmCantidates.Show
End Sub

'Displays Barack's Biography by loading it from a text file
Private Sub cmdBio_Click()

Open App.Path & "\BioTexts\obamaBio.txt" For Input As #1
Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, ObamaBio(Ctr)
    picResults.Print ObamaBio(Ctr)
    
Loop
Close #1


End Sub

'Sorts Barack's Achievements by Year by using a bubble sort
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
