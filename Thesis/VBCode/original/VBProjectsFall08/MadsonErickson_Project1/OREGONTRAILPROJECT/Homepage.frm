VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form2"
   Picture         =   "Homepage.frx":0000
   ScaleHeight     =   10665
   ScaleWidth      =   15465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "Works Cited"
      Height          =   1095
      Left            =   10560
      TabIndex        =   6
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdLearn 
      Caption         =   "Learn about Oregon Trail "
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Purchase Supplies"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdRoute 
      Caption         =   "Go on the Journey!  "
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdParty 
      Caption         =   "See who is in Your Party! "
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Caption         =   "Welcome to the great American Adventure!"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLearn_Click()
 
'Oregon Trail
'Homepage (Form 2)
'Drew Madson & Sam Erickson Oct 2008
'This is the homepage it hosts all the buttons
 
 
    MsgBox "The Oregon trail is the most righteous adventure to grace America. It is the 1800's and you and your family attempt to travel from Middle America over the Rocky Mountains to Oregon.You must hunt for food and strategically ration your suppies and rate of travel. Will you make it to Oregon before the Winter snows? Will Ma survive dyptheria? How many bison will you shoot? It is all up to you and your skills as you travel the famous and perilous OREGON TRAIL! "
    
     




End Sub

Private Sub cmdParty_Click()

    Form2.Hide
    frmpeoplewhotraveled.Show ' Bring up party search


End Sub

Private Sub cmdPurchase_Click()

    Occupation = InputBox("Are you a Farmer or French Film Maker?")
    Form2.Hide
    Form3.Show
    
    If LCase(Occupation) = "farmer" Then
        MsgBox "You shouldn't need to purchase food"
    Else
        MsgBox "While being a French Film Maker is super chic, you won't be able to buy camera equipment here."
    
    End If
End Sub
    
    
Private Sub cmdQuit_Click()

    End

End Sub

Private Sub cmdRoute_Click()
    
    MsgBox "YeHaw!"
    
    Form2.Hide
    Form4.Show
    
    

End Sub

Private Sub cmdWorksCited_Click()
Form2.Hide
Form5.Show

End Sub
