VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdParty 
      Caption         =   "Name Your Party"
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
      Left            =   480
      TabIndex        =   4
      Top             =   4440
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
      Left            =   480
      TabIndex        =   3
      Top             =   5880
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
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
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
      Left            =   480
      TabIndex        =   1
      Top             =   1440
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLearn_Click()
    
     
    'informs user about Oregon tral
    MsgBox "The Oregon trail is the most righteous adventure to grace America. It is the 1800's and you and your family attempt to travel from Middle America over the Rocky Mountains to Oregon.You must hunt for food and strategically ration your suppies and rate of travel. Will you make it to Oregon before the Winter snows? Will Ma survive dyptheria? How many bison will you shoot? It is all up to you and your skills as you travel the famous and perilous OREGON TRAIL! "
    
     




End Sub

Private Sub cmdPurchase_Click()

    Form2.Hide
    Form3.Show

End Sub
    
    
Private Sub cmdRoute_Click()
    
    MsgBox "YeHaw!"
    
    Form2.Hide
    Form4.Show
    
    

End Sub
