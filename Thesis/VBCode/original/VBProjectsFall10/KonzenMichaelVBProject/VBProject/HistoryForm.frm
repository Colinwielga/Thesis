VERSION 5.00
Begin VB.Form frmhistory 
   Caption         =   "History"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   Picture         =   "History Form.frx":0000
   ScaleHeight     =   11130
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults2 
      Height          =   3135
      Left            =   2640
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   7800
      Width           =   3735
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      Height          =   6255
      Left            =   360
      ScaleHeight     =   6195
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   1320
      Width           =   7815
   End
   Begin VB.CommandButton cmdmeetfounder 
      Caption         =   "Meet the President"
      DisabledPicture =   "History Form.frx":39973
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdshowhistory 
      Caption         =   "Show History"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this button clears the boxes
Private Sub cmdclear_Click()
picresults2 = Nothing
picresults.Cls

End Sub
'this button hides this frame and shows the main screen
Private Sub cmdgoback_Click()
frmhistory.Hide
frmmainscreen.Show
End Sub

Private Sub cmdmeetfounder_Click()
Dim meetowner As String

Open App.Path & "\MeettheOwner.txt" For Input As #2 'the file is opened
Do While Not EOF(2)
Input #2, meetowner 'i have defined my variable
    picresults.Print meetowner
Loop
Close #2 'closing the file
picresults2.Picture = LoadPicture(App.Path + "\danawhite.jpg") 'i want a picture of the president to show
End Sub

Private Sub cmdshowhistory_Click()
Dim history As String

Open App.Path & "\History.txt" For Input As #1 'i opened the file to read
Do While Not EOF(1)
Input #1, history 'the contents of the file are being defined
picresults.Print history 'printing file
Loop
    
    
Close #1 'closeing file
End Sub

