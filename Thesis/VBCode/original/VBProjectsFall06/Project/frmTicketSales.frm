VERSION 5.00
Begin VB.Form frmTicketSales 
   BackColor       =   &H00000080&
   Caption         =   "Ticket Sales"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   2280
      TabIndex        =   30
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Left            =   360
      TabIndex        =   29
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdDenver 
      BackColor       =   &H0000FFFF&
      Caption         =   "Denver"
      Height          =   375
      Index           =   6
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdUND 
      BackColor       =   &H0000FFFF&
      Caption         =   "North Dakota"
      Height          =   375
      Index           =   5
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdMariucci 
      BackColor       =   &H0000FFFF&
      Caption         =   "Mariucci Classic"
      Height          =   375
      Index           =   4
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdMSU 
      BackColor       =   &H0000FFFF&
      Caption         =   "Minnestota State"
      Height          =   375
      Index           =   3
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdMich 
      BackColor       =   &H0000FFFF&
      Caption         =   "Michigan"
      Height          =   375
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdMichST 
      BackColor       =   &H0000FFFF&
      Caption         =   "Michigan State"
      Height          =   375
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdWisc 
      BackColor       =   &H0000FFFF&
      Caption         =   "Wisconsin"
      Height          =   375
      Index           =   1
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdSCSU 
      BackColor       =   &H0000FFFF&
      Caption         =   "St. Cloud State"
      Height          =   375
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   7
      Left            =   240
      Picture         =   "frmTicketSales.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   4680
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   6
      Left            =   7200
      Picture         =   "frmTicketSales.frx":0516
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   4680
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   5
      Left            =   3600
      Picture         =   "frmTicketSales.frx":1224
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   4
      Left            =   7200
      Picture         =   "frmTicketSales.frx":1461
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   9
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   3
      Left            =   7200
      Picture         =   "frmTicketSales.frx":2125
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   2
      Left            =   3600
      Picture         =   "frmTicketSales.frx":2AFE
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   1
      Left            =   240
      Picture         =   "frmTicketSales.frx":34AC
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   0
      Left            =   3600
      Picture         =   "frmTicketSales.frx":3E1F
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCC 
      BackColor       =   &H0000FFFF&
      Caption         =   "Collorado College"
      Height          =   375
      Index           =   0
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      Picture         =   "frmTicketSales.frx":4B3D
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Jan. 26-27 7:00 PM"
      Height          =   615
      Index           =   8
      Left            =   8760
      TabIndex        =   28
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Jan. 19-20 7:00 PM"
      Height          =   615
      Index           =   7
      Left            =   5160
      TabIndex        =   27
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Dec. 29-30 7:00 PM"
      Height          =   615
      Index           =   6
      Left            =   1800
      TabIndex        =   26
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Dec. 1-2   7:00 PM"
      Height          =   615
      Index           =   5
      Left            =   8760
      TabIndex        =   25
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Nov. 25   7:00 PM"
      Height          =   615
      Index           =   4
      Left            =   5160
      TabIndex        =   24
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Nov. 24   7:00 PM"
      Height          =   615
      Index           =   3
      Left            =   1800
      TabIndex        =   23
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date:  Nov. 18-19  7:00 PM"
      Height          =   615
      Index           =   2
      Left            =   8760
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date:  Nov. 10-11  7:00 PM"
      Height          =   615
      Index           =   1
      Left            =   5160
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblCC 
      Caption         =   "Event Date: Oct 27-28 7:00 PM"
      Height          =   615
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      Caption         =   "Please pick from the following events or items:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4845
   End
   Begin VB.Label lblBuy 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Buy Gopher Hockey Tickets!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   6045
   End
End
Attribute VB_Name = "frmTicketSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmTicketSales
'Cole and John
'10/30/06
'Objective: The objective of this form is to present to user with a choice
'on which game they would like to purchase tickets for. The user can select the
'desired game by selecting the yellow command button associated with the opponent.
'The button then directs the user to the ticket purchase page.
Option Explicit

Private Sub cmdBack_Click()
    frmTickets.Visible = True       'allows user to go back
    frmTicketSales.Visible = False
End Sub

Private Sub cmdCC_Click(Index As Integer)
    frmTicketSales.Visible = False      'accesses purchase tickets page
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdDenver_Click(Index As Integer)
    frmTicketSales.Visible = False      'same as above
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdHome_Click()
    frmMain.Visible = True
    frmTicketSales.Visible = False
End Sub

Private Sub cmdMariucci_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdMich_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdMichST_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdMSU_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdSCSU_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdUND_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

Private Sub cmdWisc_Click(Index As Integer)
    frmTicketSales.Visible = False
    frmTicketPurchase.Visible = True
End Sub

