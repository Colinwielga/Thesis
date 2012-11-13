VERSION 5.00
Begin VB.Form formlocations 
   BackColor       =   &H80000007&
   Caption         =   "Number of Locations Available"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form2"
   ScaleHeight     =   8055
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdho 
      Caption         =   "BackTo Home"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   6960
      Width           =   1695
   End
   Begin VB.PictureBox piccities 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H8000000A&
      Height          =   615
      Left            =   3240
      ScaleHeight     =   615
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   7320
      Width           =   6255
   End
   Begin VB.PictureBox picmn 
      Enabled         =   0   'False
      Height          =   7815
      Left            =   3240
      Picture         =   "formlocations.frx":0000
      ScaleHeight     =   7755
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton cmdmn 
      Caption         =   "Find Your City Or a City Near You"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   840
      TabIndex        =   0
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   600
      Picture         =   "formlocations.frx":A3A0
      ScaleHeight     =   1155
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
End
Attribute VB_Name = "formlocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdho_Click()
formlocations.Hide
Form1.Show
End Sub

'by: Heather Parker
'Form name = formlocations
'form File name = Project One\formlocations.frm
' Purpose of form = Have the user tell the computer where they would like
'a Bremer Bank to be located and have the computer tell them if one exists there


Private Sub Form_Load()
Dim path As String
path = "N:\CS130\Parker, Heather\Project One\"
End Sub

Private Sub cmdmn_Click()

Dim cities(1 To 600) As String
Dim position As Integer
Dim found As Boolean
Dim A As String
Dim Ctr As Integer
Ctr = 1
Open path & "cities.txt" For Input As #1
picmn.Visible = True
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, cities(Ctr)
Loop
Close #1
A = InputBox("Enter Your City of Residence to Find if there is a Bremer Bank Near You")
found = False
position = 0
Do While Not found And (position < Ctr)
    position = position + 1
    If cities(position) = A Then
        piccities.Print "Yes,"; Tab(6); cities(position); Tab(18); "Has a Bremer Bank, Apply Today!"
        found = True
    End If
Loop
If Not found Then
    piccities.Print "There is not a Bremer Bank in Your City, Check the map to Find a City Near You"
End If
End Sub

