VERSION 5.00
Begin VB.Form frmHomeForm 
   BackColor       =   &H0080FF80&
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210.526
   ScaleMode       =   0  'User
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdPriceStart 
      BackColor       =   &H008080FF&
      Caption         =   "Price your vacation here today!!!"
      Height          =   1335
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox picCabin 
      Height          =   5100
      Left            =   360
      Picture         =   "frmHomeForm.frx":0000
      ScaleHeight     =   8621.053
      ScaleMode       =   0  'User
      ScaleWidth      =   8087.318
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.CommandButton cmdSweet 
      BackColor       =   &H008080FF&
      Caption         =   "Sweet Facts===>"
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H008080FF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblTypical 
      BackColor       =   &H8000000E&
      Caption         =   "<=== Our    Typical     Lakeside      Cabin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Left            =   6600
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "frmHomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Big Sky Resort
'frmHomeForm
'Ryan Hoffmann and Jamison Murphy
'Written on March 19, 2009
'This form was created to give the user some basic facts
'about why they may want to consider staying with us


Option Explicit
'This command goes back to previous form
Private Sub cmdBack_Click()
    frmEntryForm.Show
    frmHomeForm.Hide
End Sub

Private Sub cmdSweet_Click()
'Declaring Variables
Dim Fact(1 To 10) As String, CTR As Integer

'Here we open the file from where we get our facts from
Open App.Path & "\Facts.txt" For Input As #1

'This loads the data into an array
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Fact(CTR)
Loop

'This randomly displays the data in a picture box
Close #1
picResults.Cls
picResults.Print Fact(CInt(Int((7 * Rnd()) + 1)))

End Sub

'Moves onto the next form
Private Sub cmdPriceStart_Click()
    frmPricingForm1.Show
    frmHomeForm.Hide
End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

