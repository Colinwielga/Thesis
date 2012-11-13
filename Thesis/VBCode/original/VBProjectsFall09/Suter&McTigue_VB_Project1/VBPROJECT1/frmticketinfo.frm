VERSION 5.00
Begin VB.Form frmticketinfo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11475
   ClientLeft      =   1140
   ClientTop       =   2205
   ClientWidth     =   16965
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmticketinfo.frx":0000
   ScaleHeight     =   11475
   ScaleWidth      =   16965
   Begin VB.PictureBox picresults2 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   8040
      ScaleHeight     =   2355
      ScaleWidth      =   8475
      TabIndex        =   14
      Top             =   2040
      Width           =   8535
   End
   Begin VB.CommandButton cmdschedule 
      BackColor       =   &H00000000&
      Caption         =   "Get the United States first round schedule"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtgame 
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdreturntoform 
      Caption         =   "Return to main page"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   10
      Top             =   10680
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Start Over"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   9
      Top             =   10680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdprintticket 
      BackColor       =   &H00404040&
      Caption         =   "Print Ticket"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   8
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtlastname 
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox txtmiddleinitial 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtfirstname 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   4080
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label lblgame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your desired game:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lbllastname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your last name here:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label lblmiddleinitial 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your middle initial here:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblfirstname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your first name here:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblticketinfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ticket Information"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmticketinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'USAsoccer
'frmticketinfo
'author Marty and Sean
'October 13
'this is meant to help find scheduled games, and purchase tickets and set up ID
Private Sub cmdclear_Click()
    'this cmd button will erase the ticket in the picture box
    picresults.Cls
    
End Sub

Private Sub cmdprintticket_Click()
    'this cmd button will print the information needed for a ticket gathered from the text boxes
    
    Dim lastname As String, middleinitial As String, firstname As String
    Dim opponent As String
    
    'declare variables
    lastname = txtlastname.Text
    opponent = txtgame.Text
    firstname = txtfirstname.Text
    middleinitial = txtmiddleinitial.Text
    
    'print results (ticket)
    picresults.Print "Your World Cup ID"
    picresults.Print "_______________________________________________________"
    picresults.Print Left(firstname, 1); " "; Left(middleinitial, 1); " "; lastname
    picresults.Print ""
    picresults.Print "Opponent:"
    picresults.Print "__________________________________________________"
    picresults.Print opponent
    picresults.Print "__________________________________________________"
    'This shows the Current date
    picresults.Print "The date of purchase was "; Date
    
    'this will show the clear button(start over)
    cmdclear.Visible = True
    
End Sub

Private Sub cmdreturntoform_Click()
    'this cmd button will return the user to the main page
    frmticketinfo.Hide
    Form1.Show

End Sub

Private Sub cmdschedule_Click()
    'this cmd button will print the first round schedule for the United States' team
    'in the picture box named picresults2
    
    picresults2.Print ""
    picresults2.Print "Opponent"; Tab(20); "Date"; Tab(30); "Time"; Tab(55); "Venue"; Tab(85); "Location"
    picresults2.Print "*****************************************************************************************************************************************************************************************************"
    picresults2.Print ""
    picresults2.Print "Germany"; Tab(20); "12 June"; Tab(30); "2:00 pm"; Tab(55); "Moses Mabhida Stadium"; Tab(85); "Durban"
    picresults2.Print ""
    picresults2.Print "Australia"; Tab(20); "18 June"; Tab(30); "2:00 pm"; Tab(55); "Free State Stadium"; Tab(85); "Bloemfontein (Mangaung)"
    picresults2.Print ""
    picresults2.Print "Paraguay"; Tab(20); "24 June"; Tab(30); "11:00 am"; Tab(55); "Peter Mokaba Stadium"; Tab(85); "Polokwane"
    picresults2.Print ""
End Sub
