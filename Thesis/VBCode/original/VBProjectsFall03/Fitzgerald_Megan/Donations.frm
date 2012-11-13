VERSION 5.00
Begin VB.Form frmDonations 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Donations (Megan Fitzgerald)"
   ClientHeight    =   7665
   ClientLeft      =   2790
   ClientTop       =   1950
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10140
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdReadGD 
      BackColor       =   &H00FF8080&
      Caption         =   "Click to see Goods Distributed"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   5295
      Left            =   2640
      ScaleHeight     =   5235
      ScaleWidth      =   7395
      TabIndex        =   3
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdReadPF 
      BackColor       =   &H00FF8080&
      Caption         =   "Click to see Project Funding"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Donations.frx":0000
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Donations"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmDonations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmDonations (Donations.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: To inform the user of donations that Amigos for Chirst
                        'has received as a non-profit organization and to
                        'allow them to see Project Funding and Goods Distributed.
                        
                    
Option Explicit

'"PATH" refers to "M:\CS130\MeganFitzgerald\Fitzgerald_Megan\".
'By declaring this as a form level variable it makes it easier to read the code.
Dim PATH As String

'Clicking this button will allow the user to see donations for Goods Distributed.
Private Sub cmdReadGD_Click()
'Clear the picture box for repeated use.
picResults.Cls

Dim ProjectRecipients(1 To 8) As String, Money(1 To 8) As Double
Dim I As Integer, Total As Single

'Open the file "GoodsDistributed.txt" on Channel #2
Open PATH & "GoodsDistributed.txt" For Input As #2

'Print these headings.
picResults.Print "Amigos for Christ Donation Project Destination Statement"
picResults.Print "*****************************************************************************************************"
picResults.Print "*****************************************************************************************************"

picResults.Print "Goods Distributed"; Tab(41); "Total through May 2002"
picResults.Print "******************************************************************************************************"

'Put the information from Channel #2 into an several parallel arrays
'so that this information can be used and manipulated later.
For I = 1 To 8
    Input #2, ProjectRecipients(I), Money(I)
    picResults.Print ProjectRecipients(I); Tab(41); FormatCurrency(Money(I))
    Total = Total + Money(I)
Next I

picResults.Print "----------------------------------------------------------------------------------------------------------"
picResults.Print "TOTAL GOODS DISTRIBUTED"; Tab(41); FormatCurrency(Total)

Close #2

End Sub

'Clicking this button will allow the user to see the donations for Project Funding.
Private Sub cmdReadPF_Click()

'Clear the picture box for repeated use.
picResults.Cls
Dim ProjectsFunded(1 To 10) As String, Money(1 To 10) As Single
Dim I As Integer, Total As Single

'Open the file "ProjectsFunded.txt" on Channel #1
Open PATH & "ProjectsFunded.txt" For Input As #1


'Print these headings.
picResults.Print "Amigos for Christ Donation Project Destination Statement"
picResults.Print "*****************************************************************************************************"
picResults.Print "*****************************************************************************************************"

picResults.Print "Projects Funded"; Tab(41); "Total through May 2002"
picResults.Print "********************************************************************************************************"

'Input the information on Channel #2 into several parallel arrays so that
'this information can used and manipulated later.
For I = 1 To 10
    Input #1, ProjectsFunded(I), Money(I)
    picResults.Print ProjectsFunded(I); Tab(41); FormatCurrency(Money(I))
    Total = Total + Money(I)
Next I

picResults.Print "----------------------------------------------------------------------------------------------------------"
picResults.Print "TOTAL PROJECT FUNDING"; Tab(41); FormatCurrency(Total)

Close #1

End Sub


Private Sub cmdReturn_Click()

'This makes it possible for the user to return to the "Homepage" easily.
'This will hide the "Donations" page and bring the "Homepage"
'into view.
frmDonations.Hide
frmHomepage.Show

End Sub


Private Sub Form_Load()
'Bring up this picture when this form is loaded.
picResults1.Picture = LoadPicture("N:\CS130\handin\Fitzgerald_Megan\Pictures\Kids near House.jpg")
PATH = "N:\CS130\handin\Fitzgerald_Megan\"
End Sub

