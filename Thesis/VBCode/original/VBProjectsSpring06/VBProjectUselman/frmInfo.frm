VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00008000&
   Caption         =   "Minnesota Wildlife Information"
   ClientHeight    =   8550
   ClientLeft      =   1410
   ClientTop       =   1125
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12615
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00008000&
      Height          =   2895
      Left            =   5280
      ScaleHeight     =   2835
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Go Back To Main Menu"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1800
      ScaleHeight     =   2115
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   1920
      Width           =   9615
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00008000&
      Caption         =   "Click Here To Explore Minnesota Wildlife"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblMe 
      BackColor       =   &H00008000&
      Caption         =   "Lance Uselman"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmInfo (frmInfo.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the form: This form allows the user to input a name of an animal
'                     into a inputbox and then receive information about the
'                     animal through two picture boxes.

Option Explicit

Private Sub cmdBack_Click()
    frmMain.Show
    frmInfo.Hide    'This button allows the user to go back to the main form.
End Sub

Private Sub cmdInfo_Click()
    Dim A, info, infor As String    'This button asks the user to input an animal name in an input box and then prints the associated results into the two picture boxes.
    A = InputBox("Enter a species found in Minnesota", "Enter Species")
    picResults.Cls
    picResults2.Cls
    If A = "White-tailed Deer" Then
        Open App.Path & "\deer.txt" For Input As #1
        Do Until EOF(1)
            Input #1, info
            picResults.Print info
            picResults2.Picture = LoadPicture(App.Path & "\wdeer.jpg")
        Loop
        Close #1
    ElseIf A = "Black Bear" Then
        Open App.Path & "\bear.txt" For Input As #2
        Do Until EOF(2)
            Input #2, infor
            picResults.Print infor
            picResults2.Picture = LoadPicture(App.Path & "\blackbear.jpg")
        Loop
        Close #2
    Else
        MsgBox "Chosen species is currently not in the database. (Try White-tailed Deer or Black Bear)", , "Sorry"
    End If
End Sub
