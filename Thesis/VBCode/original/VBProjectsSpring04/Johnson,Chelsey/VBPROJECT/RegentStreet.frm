VERSION 5.00
Begin VB.Form RegentStreet 
   BackColor       =   &H00404080&
   Caption         =   "Regent Street"
   ClientHeight    =   12465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   12465
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   9
      Top             =   11520
      Width           =   1455
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12600
      TabIndex        =   8
      Top             =   10440
      Width           =   2055
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Click here to learn the hours of eating services within Covent Gardens Market"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picResults2 
      Height          =   4335
      Left            =   8040
      ScaleHeight     =   4275
      ScaleWidth      =   6795
      TabIndex        =   6
      Top             =   1800
      Width           =   6855
   End
   Begin VB.PictureBox picGarden 
      Height          =   2895
      Left            =   120
      Picture         =   "RegentStreet.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   15195
      TabIndex        =   4
      Top             =   6840
      Width           =   15255
   End
   Begin VB.PictureBox picbritmuseum 
      Height          =   3015
      Left            =   120
      Picture         =   "RegentStreet.frx":B808
      ScaleHeight     =   2955
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   3360
      Width           =   7335
   End
   Begin VB.PictureBox picResults 
      Height          =   1575
      Left            =   2880
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton cmdstreet 
      Caption         =   "Click here to guess the name of a famous street in the Regent Street District of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   11880
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Regent Covenant Garden Market"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "This is the British Museum, Click on the picture to learn more about it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   7335
   End
End
Attribute VB_Name = "RegentStreet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: RegentStreet (RegentStreet.frm)
'Author: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: The purpose of this form is to inform the user of the history of 7 Dials Street and the British Museum
                    'They are also able to guess the street name by guessing the number before 'Dials'.  They are also
                    'albe to view the normal hours for the different varieties of eating places within Covent Gardens.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdload_Click()
Dim Hours As String
Dim Days(1 To 7) As String
Dim Morn(1 To 7) As Integer
Dim Night(1 To 7) As Integer
Dim J As Integer
Dim CTR As Integer
picResults2.Cls
'Getting variable from the user
Hours = InputBox("Please type in one of the following- Cafes, Resteraunts, or Bars; to learn the hours they are open.", "Food")
If Hours = "Cafes" Then
    Open "N:\CS130\handin\Johnson, Chelsey\VBPROJECT\Cafehours.txt" For Input As #1  'Opening data
    CTR = 0
    Do While Not EOF(1) 'Filling array for Cafes
        CTR = CTR + 1
        Input #1, Days(CTR), Morn(CTR), Night(CTR)
    Loop
    picResults2.Print "The average hours for Cafes in Covent Garden are"
    picResults2.Print "*********************************************************************************************"
    picResults2.Print "Day"; Tab(25); "Opening Hours"; Tab(50); "Closing Hours"
    For J = 1 To 7
        picResults2.Print ; Days(J); Tab(25); Morn(J); Tab(50); Night(J) 'Printing out data of array
    Next J
End If
Close #1
If Hours = "Resteraunts" Then
    Open "N:\CS130\handin\Johnson, Chelsey\VBPROJECT\resthours.txt" For Input As #2  'Opening data
    CTR = 0
    Do While Not EOF(2) 'Filling array for Resteraunts
        CTR = CTR + 1
        Input #2, Days(CTR), Morn(CTR), Night(CTR)
    Loop
    picResults2.Print "The average hours for Resteraunts in Covent Garden are"
    picResults2.Print "*********************************************************************************************"
    picResults2.Print "Day"; Tab(25); "Opening Hours"; Tab(50); "Closing Hours"
    For J = 1 To 7
        picResults2.Print ; Days(J); Tab(25); Morn(J); Tab(50); Night(J) 'Printing out data of array
    Next J
End If
Close #2
If Hours = "Bars" Then
 Open "N:\CS130\handin\Johnson, Chelsey\VBPROJECT\barhours.txt" For Input As #3 'Opening Data
    CTR = 0
    Do While Not EOF(3) 'Filling array for bars
        CTR = CTR + 1
        Input #3, Days(CTR), Morn(CTR), Night(CTR)
    Loop
    picResults2.Print "The average hours for Bars in Covent Garden are"
    picResults2.Print "*********************************************************************************************"
    picResults2.Print "Day"; Tab(25); "Opening Hours"; Tab(50); "Closing Hours"
    For J = 1 To 7
        picResults2.Print ; Days(J); Tab(25); Morn(J); Tab(50); Night(J) 'Printing out data of array
     Next J
End If
Close #3
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returning user back to Map of London page, so  they are able to choose a new district
RegentStreet.Hide
MapLondon.Show
End Sub

Private Sub cmdstreet_Click()
Dim N As Integer
Dim Diff As Integer
'Getting Variable from user
N = InputBox("There is a famous street named '________ Dials Street', How many dials do you think there are?", "Dials")
picResults.Cls
picResults.Print N; "Dials Street, was your guess."
picResults.Print "7 Dials Street is the actual street name."
If N > 7 Then 'Comparing user data
    Diff = N - 7
    picResults.Print "Your guess was"; Diff; "number(s) away from the answer."
End If
If N < 7 Then 'Comparing user data
    Diff = 7 - N
    picResults.Print "Your guess was"; Diff; "number(s) away from the answer."
End If
If N = 7 Then 'Comparing user data
    picResults.Print "7 dials was the correct answer, congradulations!"
End If
picResults.Print "Seven Dials was named because it is a long street that"
picResults.Print "connects seven other streets in a row."
End Sub



Private Sub picbritmuseum_Click()
'Printing out history of British Museum when the user clicks on the picture of it.
MsgBox "The origins of the British Museum lie in the will of the physician, naturalist and collector, Sir Hans Sloane (1660-1753).The Museum was first housed in a 17th-century mansion, Montagu House, in Bloomsbury on the site of today's building. On 15 January 1759 the British Museum opened to the public. With the exception of two World Wars, when parts of the collection were evacuated, it has remained open ever since.The Museum celebrated its 250th anniversary in 2003. ", , "British Museum"
End Sub
