VERSION 5.00
Begin VB.Form frmberlin 
   AutoRedraw      =   -1  'True
   Caption         =   "Berlin"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Display"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   20
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load "
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdquitber 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdwall 
      Caption         =   "What is Berlin Wall??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   4815
   End
   Begin VB.CommandButton cmdrank 
      Caption         =   "See the Ranking"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdwall_height 
      Caption         =   "How big- get result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdwall_long 
      Caption         =   "How long- get result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtwall_height 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox txtwall_long 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton cmdimg 
      Caption         =   "Get Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtimmg 
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdanswer 
      Caption         =   "Get Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtrate 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdpicture 
      BackColor       =   &H8000000D&
      Caption         =   "What are these pictures??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFF80&
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picoutput 
      Height          =   4215
      Left            =   5640
      ScaleHeight     =   4155
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   2880
      Width           =   4935
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Cmdquiz 
      Caption         =   "Show quiz"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblwall 
      BackColor       =   &H80000002&
      Caption         =   "How big(feet) and long (mile) the Berlin Wall was which was falled in 1989 after the German reunification ??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label lblimg 
      BackColor       =   &H80000002&
      Caption         =   "Germany has many immigrations like the US. Where do they mainly come from?             Hint) East Europe, Near East etc    "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblrate 
      BackColor       =   &H80000002&
      Caption         =   "What is the current unemployment rate in Germany? "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "BERLIN, GERMANY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   5640
      Picture         =   "frmberlin.frx":0000
      Top             =   120
      Width           =   5115
   End
   Begin VB.Image Image4 
      Height          =   3780
      Left            =   -120
      Picture         =   "frmberlin.frx":38442
      Top             =   -120
      Width           =   7500
   End
   Begin VB.Image Image3 
      Height          =   4050
      Left            =   5160
      Picture         =   "frmberlin.frx":94914
      Top             =   3120
      Width           =   7275
   End
   Begin VB.Image Image2 
      Height          =   4050
      Left            =   0
      Picture         =   "frmberlin.frx":F48F6
      Top             =   3480
      Width           =   5190
   End
End
Attribute VB_Name = "frmberlin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Single
Dim nation(1 To 20) As String
Dim popul(1 To 20) As Single
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Berlin (frmberlin.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'This form introduce Berlin and provide some trivia questions
'about the city.


Private Sub cmdanswer_Click()
'This part ask user to inout their estimate of Germany's jobless
'rate, and the program determine how close their estimate is to
'the actual current unemplyment rate of Germany, by using if
'statement.
Dim unemp As Single
unemp = txtrate.Text

Select Case unemp
    Case Is <= 2
        picoutput.Print "your anwer is too low."
    Case 2 To 4.8
        picoutput.Print "Still too low...it's almost Japan's current unemplyment rate"
    Case 4.9 To 6
        picoutput.Print "...the same as current rate in the US"
    Case 6 To 9
        picoutput.Print "It's close... other EU nationas are like this"
    Case 9 To 12
        picoutput.Print "Exactly!! the Current Unemployment rate of Germany is 11.15"
        picoutput.Print "according to the recent edition of Economist"
    
    Case 12 To 15
        picoutput.Print "That's too high...their rate is not that bad"
        
    Case Else
        picoutput.Print "Try again...the answer will be below 15"
    
    End Select
 
    
End Sub


Private Sub cmdclear_Click()
'This code works to clear the picture box.
picoutput.Cls

End Sub

Private Sub cmdimg_Click()
'The user answer the quiz about the foreigner in Germany. The user
'input the nation they think is the major immirtant origin of
'germany and the program tells how many foreignes from each nation
'there are in the country from the data in the array.

Dim answer As String
answer = txtimmg.Text
X = 1
Do Until answer = nation(X)
    X = X + 1

Loop
picoutput.Print nation(X), popul(X)


    

End Sub

Private Sub cmdload_Click()
'This code works to load the file for a quiz about immigration
'in Germany.
 
Open App.Path & "\immg.txt" For Input As #1
For X = 1 To 20
    Input #1, nation(X), popul(X)
Next X

End Sub

Private Sub cmdmain_Click()
'This botton take users back to the starting page.
frmberlin.Hide
frmmain.Show
End Sub

Private Sub cmdpicture_Click()
'This code works to hide the labels, bottons, and boxes and show
'the pictures on the back. Also, this code provides the description
'of each pictures on the picturebox.
cmdpicture.Visible = False
cmdmain.Visible = False
cmdquitber.Visible = False
lblrate.Visible = False
lblimg.Visible = False
lblwall.Visible = False
txtrate.Visible = False
txtimmg.Visible = False
txtwall_long.Visible = False
txtwall_height.Visible = False
cmdanswer.Visible = False
cmdimg.Visible = False
cmdload.Visible = False
cmdclear.Visible = False
cmdrank.Visible = False
cmdwall.Visible = False
cmdwall_long.Visible = False
cmdwall_height.Visible = False

picoutput.Cls
picoutput.Print "From above left"
picoutput.Print "German Budentag (Congress)"
picoutput.Print "Alexnader Platz and Fernsehturm (TV Tower)"
picoutput.Print "Brandenburg Gate"
picoutput.Print "Fall of Berlin Wall in 1989"
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdquitber_Click()
End
End Sub

Private Sub Cmdquiz_Click()
'This code shows all boxes, labels and buttons again to go over
'the quizs.
picoutput.Cls
'make all buttons, boxes, labels visible again.
lblrate.Visible = True
lblimg.Visible = True
lblwall.Visible = True
txtrate.Visible = True
txtimmg.Visible = True
txtwall_long.Visible = True
txtwall_height.Visible = True
cmdanswer.Visible = True
cmdimg.Visible = True
cmdrank.Visible = True
cmdwall.Visible = True
cmdwall_long.Visible = True
cmdwall_height.Visible = True
cmdpicture.Visible = True
cmdmain.Visible = True
cmdclear.Visible = True
cmdload.Visible = True
cmdquitber.Visible = True


End Sub

Private Sub cmdrank_Click()
'This code works to show users the overall data in the array.
For X = 1 To 20
picoutput.Print nation(X), popul(X)
Next X




End Sub

Private Sub cmdwall_Click()
'This button take user to the detail explanation of Berlin Wall,
' which is on the new form/
frmwall.Show

End Sub

Private Sub cmdwall_height_Click()
'This code determine how close the height of
'Berlin wall he/she input to the actual.
Dim how_big As Single
how_big = txtwall_height.Text

Select Case how_big
    Case Is <= 3
        picoutput.Print "Too Small... Try Again."
    Case 4 To 6
        picoutput.Print "The wall is a little bit more bigger."
    Case 7 To 8
        picoutput.Print "You are almost getting correct answer..."
    Case 9 To 10
        picoutput.Print "Correct! The height of Berlin wall was 11.81 feet high(3.5m)"
    Case Else
        picoutput.Print "Think carefully again...the wall wasn't that big"

End Select

End Sub



Private Sub cmdwall_long_Click()
'This code determine how close the lenth of
'Berlin wall he/she input to the actual size.
Dim how_long As Single
how_long = txtwall_long.Text

Select Case how_long
    Case Is <= 50
        picoutput.Print "Too Short... Try Again."
    Case 51 To 80
        picoutput.Print "The wall is a little bit more longer."
    Case 81 To 92
        picoutput.Print "You are almost getting correct answer..."
    Case 93 To 100
        picoutput.Print "Correct! The total lenght around the wall was 96 mile(155km)"
    Case Else
        picoutput.Print "Think carefully again"

End Select

End Sub

