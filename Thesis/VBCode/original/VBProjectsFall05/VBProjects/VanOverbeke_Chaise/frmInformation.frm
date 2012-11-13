VERSION 5.00
Begin VB.Form frmInformation 
   BackColor       =   &H00008000&
   Caption         =   "Further Information"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrowth 
      Caption         =   "Company Objectives and Future Goals"
      Height          =   1095
      Left            =   7920
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdFounder 
      Caption         =   "About the founder of Chaise Van Air"
      Height          =   1095
      Left            =   5400
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox picOutbox 
      Height          =   6735
      Left            =   360
      ScaleHeight     =   6675
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton cmdMainMenu4 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   8040
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblMembername 
      BackColor       =   &H00008000&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label lblName 
      BackColor       =   &H00008000&
      Caption         =   "By: Chaise VanOverbeke"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   3585
      Left            =   4920
      Picture         =   "frmInformation.frx":0000
      Top             =   2160
      Width           =   5280
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   5160
      Picture         =   "frmInformation.frx":3DA22
      Top             =   6000
      Width           =   1740
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Airline Option(Project1.vbp)
'Form Name : frmInformation(frmInformation.frm)
'Author: Chaise VanOverbeke
'Date : Wednesday October 26, 2005

'Purpose of this form:  If the user wishes to find out more information about the
                        'the company, they can visit this form, which allows the
                        'user to find out more information about the founder of the
                        'company and what direction the company is heading and its
                        'goals and objectives.

Option Explicit

Private Sub cmdFounder_Click()
    picOutbox.Cls   'clears any information that exists in the picture box.
    picOutbox.Print     'prints a blank line in the picture box.
    picOutbox.Picture = LoadPicture(App.Path & "\Chaise.bmp")   'loads the previous file (a picture) into the picture box.
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print
    picOutbox.Print "     The Chaise Van Air flight company was founded in"     'the rest of this command button displays various information about the founder in the picture box in paragraph form.
    picOutbox.Print "Septemberof 2004 and began business in April of 2005. The"
    picOutbox.Print "company was founded by the brilliant entrepreneur, Chaise"
    picOutbox.Print "VanOverbeke, a man full of visions.  VanOverbeke graduated"
    picOutbox.Print "from St. John's University in 1988 and went on to get his"
    picOutbox.Print "Masters in Business Management to years later from the"
    picOutbox.Print "University of Minnesota."
    picOutbox.Print
    picOutbox.Print "     If you have any questions or comments for Chaise"
    picOutbox.Print "VanOverbeke you can e-mail him at crvanoverbe@csbsju.edu,"
    picOutbox.Print "and you are guarenteed to have a response from him within"
    picOutbox.Print "two weeks time.  Chaise VanOverbeke appreciates your"
    picOutbox.Print "business and loyalty; it will not be overlooked."
End Sub

Private Sub cmdGrowth_Click()
    picOutbox.Cls   'clears any information that exists in the picture box.
    picOutbox.Picture = LoadPicture("")    'tells the program not to load a picture into the picture box.
    picOutbox.Print "Chaise Van Air is a growing company with many objectives"  'the rest of this command button displays various information in the picture box in paragraph form.
    picOutbox.Print "and many goals for the future..."
    picOutbox.Print
    picOutbox.Print "     Unlike other companies, we make you, the customer our"
    picOutbox.Print "number one priority and we take pride in our excellent"
    picOutbox.Print "service."
    picOutbox.Print
    picOutbox.Print "     We are constantly offering monthly deals and various"
    picOutbox.Print "drawings for excellent prices."
    picOutbox.Print
    picOutbox.Print "     Our work environment is based on a teamwork mentality"
    picOutbox.Print "and equality where everyone's job is important. We seek to"
    picOutbox.Print "evenly spread out our wealth distribution more so than our"
    picOutbox.Print "competitors to better limit the amount of employee strikes"
    picOutbox.Print "therefore maintaing a solid amount of business and validity"
    picOutbox.Print "for you the customer."
    picOutbox.Print
    picOutbox.Print "     We are a new company that is growing in an industry"
    picOutbox.Print "where other companies are dying out, because we hold the"
    picOutbox.Print "competitive edge. Chaise Van Air currently has flight"
    picOutbox.Print "routes to 20 of the 52 states in the U.S, and by Feb of"
    picOutbox.Print "2006, Chaise Van Air will have routes to 36 states.  Our"
    picOutbox.Print "big goal in progress at the moment is going global, which"
    picOutbox.Print "We would have flights to Europe, Asia, and South America,"
    picOutbox.Print "by the middle of 2008."
End Sub

Private Sub cmdMainMenu4_Click()
    frmInformation.Hide
    frmMainForm1.Show      'goes back to the main menu, so the user can visit the other forms.
End Sub

