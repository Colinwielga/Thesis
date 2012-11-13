VERSION 5.00
Begin VB.Form frmMilage 
   BackColor       =   &H00000000&
   Caption         =   "Milage"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtOption 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblOption2 
      BackColor       =   &H00000000&
      Caption         =   "I need my groceries delivered to number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   3120
      Width           =   7095
   End
   Begin VB.Label lblOptions 
      BackColor       =   &H00000000&
      Caption         =   "#1 - Groceries need to be delivered to CSB #2 - Groceries need to be delivered to SJU   #3 - Groceries need to be delivered to SCS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblDelivery 
      BackColor       =   &H00000000&
      Caption         =   "Delivery Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmMilage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmMilage.Hide
frmCheckOut.Show
'direct user to CheckOut form
End Sub

Private Sub cmdNext_Click()
milage = txtOption.Text
frmMilage.Hide
frmConfirmation.Show
'retrieve milage option from user, then direct to Confirmation form
End Sub

