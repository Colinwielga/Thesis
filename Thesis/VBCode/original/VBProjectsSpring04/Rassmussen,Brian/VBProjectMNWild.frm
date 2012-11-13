VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   Picture         =   "VB Project MN Wild.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Exit Your Wild Adventure"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   4440
      Width           =   2175
   End
   Begin VB.PictureBox Picplayer 
      Height          =   3015
      Left            =   6120
      ScaleHeight     =   2955
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin VB.PictureBox Picstats 
      Height          =   1935
      Left            =   6600
      ScaleHeight     =   1875
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Be A Fan Club Member For The Minnesota Wild"
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Cmdtotal 
      Caption         =   "Click Here to see which Wild Player has the most Total Points"
      Height          =   975
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton CmdSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "VB Project MN Wild.frx":A082
      Left            =   120
      List            =   "VB Project MN Wild.frx":A095
      TabIndex        =   1
      Text            =   "Choose your favorite Wild player out of their top five stars"
      Top             =   5160
      Width           =   4695
   End
   Begin VB.CommandButton CmdWebsite 
      Caption         =   "View Your Wild Web Site"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   2640
      Picture         =   "VB Project MN Wild.frx":A0E8
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackColor       =   &H00004000&
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Click submit to view their Statistics and picture "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Wild Hockey(Brian Rassmussen's Wild Hockey Program.vbp)
'Author: Brian Rassmussen
'Date Written: March 13, 2004
'Purpose of the Project: 'To introduce people to the state of hockey and the Minnesota Wild hockey team.  Gives them a chance to view the top five players on the wild and there pictures and statistics
                     
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim J As Integer, Player(1 To 5) As String, Goals(1 To 5) As String, Assists(1 To 5) As String, Points(1 To 5) As String
Dim PATH As String

'The following code I found on the Microsoft VB news group
' Thanks to g@microsoft.com fpr the following code
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Private Sub CmdQuit_Click()
End
End Sub

Private Sub CmdSubmit_Click()


Select Case Combo1.ListIndex
'This combo box will allow the user to pick their favorite player out of the list and view there picture as well as goals, assists, and total points

Case 0
'code for Gaborik's Picture and Statistics
Picplayer.Picture = LoadPicture(PATH & "Images\Gaborik.jpg")
Picstats.Cls
Picstats.Print "Name:", Player(1)
Picstats.Print "Goals:", Goals(1)
Picstats.Print "Assists:", Assists(1)
Picstats.Print "Total Points:", Points(1)

Case 1
'Code for Richard Park's Picture and Statistics
Picplayer.Picture = LoadPicture(PATH & "Images\Park.jpg")
Picstats.Cls
Picstats.Print "Name:", Player(2)
Picstats.Print "Goals:", Goals(2)
Picstats.Print "Assists:", Assists(2)
Picstats.Print "Total Points:", Points(2)


Case 2
'Code for Annti Laaksonen's Picture and Statistics
Picplayer.Picture = LoadPicture(PATH & "Images\Annti.jpg")
Picstats.Cls
Picstats.Print "Name:", Player(3)
Picstats.Print "Goals:", Goals(3)
Picstats.Print "Assists:", Assists(3)
Picstats.Print "Total Points:", Points(3)
Case 3
'Code for Pierre-Marc Bouchard's Picture and Statistics
Picplayer.Picture = LoadPicture(PATH & "Images\Bouchard.jpg")
Picstats.Cls
Picstats.Print "Name:", Player(4)
Picstats.Print "Goals:", Goals(4)
Picstats.Print "Assists:", Assists(4)
Picstats.Print "Total Points:", Points(4)
Case 4
'Code for Wes Walz's Picture and Statistics
Picplayer.Picture = LoadPicture(PATH & "Images\Walz.jpg")
Picstats.Cls
Picstats.Print "Name:", Player(5)
Picstats.Print "Goals:", Goals(5)
Picstats.Print "Assists:", Assists(5)
Picstats.Print "Total Points:", Points(5)
Case Else
End Select
End Sub

Private Sub Cmdtotal_Click()
'This button will find player with most points and print stats
Dim N As Integer, M As Integer
N = 0
M = 0
For J = 1 To 5
    If Points(J) > N Then
    N = Points(J)
    M = J
End If
Next J
Picstats.Cls
Picstats.Print "Name:", Player(M)
Picstats.Print "Goals:", Goals(M)
Picstats.Print "Assists:", Assists(M)
Picstats.Print "Total Points:", Points(M)

Select Case M
'Select Case will print the picture of the player with the most points
Case 1
Picplayer.Picture = LoadPicture(PATH & "Images\Gaborik.jpg")
Case 2
Picplayer.Picture = LoadPicture(PATH & "Images\Park.jpg")
Case 3
Picplayer.Picture = LoadPicture(PATH & "Images\Annti.jpg")
Case 4
Picplayer.Picture = LoadPicture(PATH & "Images\Bouchard.jpg")
Case 5
Picplayer.Picture = LoadPicture(PATH & "Images\Walz.jpg")
End Select

End Sub

Private Sub CmdWebsite_Click()
'The following code I found on the Microsoft VB news group
' Thanks to g@microsoft.com fpr the following code


   Dim r As Long
   r = ShellExecute(0, "open", "http://www.wild.com", 0, 0, 1)

End Sub



Private Sub Command4_Click()
Dim r As Long
   r = ShellExecute(0, "open", "http://wild.com/fans/002/511/", 0, 0, 1)

End Sub



Private Sub Form_Load()
PATH = "N:\CS130\Handin\Rassmussen, Brian\"
Open PATH & "players.txt" For Input As #1
For J = 1 To 5
    Input #1, Player(J), Goals(J), Assists(J), Points(J)
Next J
Close #1
End Sub
