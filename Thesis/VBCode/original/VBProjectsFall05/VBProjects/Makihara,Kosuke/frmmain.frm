VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Main"
   ClientHeight    =   5595
   ClientLeft      =   630
   ClientTop       =   165
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5040
      Width           =   3975
   End
   Begin VB.CommandButton cmdquit1 
      Caption         =   "No Passport?"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdsin 
      Caption         =   "Go to Singapore"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton cmdberlin 
      Caption         =   "Go to Berlin, Germany"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdtokyo 
      Caption         =   "Go to Tokyo, Japan"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaskColor       =   &H0080FF80&
      TabIndex        =   0
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Image imgSin 
      Height          =   2925
      Left            =   0
      Picture         =   "frmmain.frx":0000
      Top             =   2640
      Width           =   3915
   End
   Begin VB.Image imagBerlin 
      Height          =   2820
      Left            =   3840
      Picture         =   "frmmain.frx":25572
      Top             =   0
      Width           =   3915
   End
   Begin VB.Image imgTokyo 
      Height          =   2670
      Left            =   0
      Picture         =   "frmmain.frx":49574
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Main (frmmain.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'
'This project introduce three cities in the world I went this summer: Tokyo,
'which the capital of my home country, Berlin, which I had a summer class for 2 months,
'and Singapore, which I always transit on the way to Asia, Europe and Oceania.
'In each page, one encounter some quizs, a kind of trivia, about each city and
'country, such as the GDP of the nation.
'
'The first page provide the selection of the cities one want to take a look.
'By clicking each botton, one can go to the each city's form and can do the activity.,
'

Private Sub cmdberlin_Click()
'By pushing the command button of the city, one can go to the new form which
'introduce the city.
frmberlin.Show

End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdquit1_Click()
'If one doesn't have passport and click this botton, called "No Passport",
'a messeage in below is shown in the message box.

MsgBox "Sorry...YOU NEED YOUR PASSPORT TO LEAVE THE US!! GET ONE AND COME BACK LATER~~", , "Goodbye~"


End Sub

Private Sub cmdsin_Click()
'By pushing the command button of the city, one can go to the new form which
'introduce the city.
frmsingapore.Show

End Sub

Private Sub cmdtokyo_Click()
'By pushing the command button of the city, one can go to the new form which
'introduce the city.
frmtokyo.Show


End Sub
