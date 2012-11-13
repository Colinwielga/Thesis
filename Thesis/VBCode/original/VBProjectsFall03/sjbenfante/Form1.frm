VERSION 5.00
Begin VB.Form HomePage 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History of the Greeen Bay Packers!"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdTicketPrices 
      Caption         =   "Do you want any tickets?  How about how much they are?  Just Click HERE!"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHistory_Click()
HomePage.Hide
WhoCoachedWhen.Show
'this hides the first form and shows the fourth form'
End Sub

Private Sub cmdQuit_Click()
    End
'this automatically end the program'
End Sub

Private Sub cmdTicketPrices_Click()
HomePage.Hide
TicketPricing.Show
'this hides the first form and shows the second form'
End Sub

Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
'this allows pictures and images to be put into different files along with the program'
End Sub
