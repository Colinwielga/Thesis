VERSION 5.00
Begin VB.Form frmEnd 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      BackColor       =   &H0000FFFF&
      Caption         =   "End"
      Height          =   735
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdfinaltotal 
      BackColor       =   &H0000FFFF&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Brush Script Std"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF00&
      Height          =   3615
      Left            =   480
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   3840
      Picture         =   "frmEnd.frx":0000
      Top             =   1080
      Width           =   6000
   End
   Begin VB.Label lblreceipt 
      BackColor       =   &H00FFFF00&
      Caption         =   "Your Receipt:"
      BeginProperty Font 
         Name            =   "GulimChe"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmEnd
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/9/08
'Objective: This is the ending page in our program
'We print out a receipt for the user's entire trip.


Private Sub cmdend_Click()

'Here the user finishes the program a message box pops up to thank them for using our company

MsgBox ("Thank you for planning your vacation with us. We hope to see you again. Have fun on your trip!!")

End

End Sub

Private Sub cmdfinaltotal_Click()

'Here we gather the totals from the activites, hotel, and car rental forms
'We add them together to get a trip total
'We display all these total in a picture box in a receipt format
'Only after printing out their receipt can they exit the program

Finaltotal = runningtotal + carRentaltotal + Hoteltotal

picResults.Print "Hotel Expense:", FormatCurrency(Hoteltotal)
picResults.Print "Activities Expense:", FormatCurrency(runningtotal)
picResults.Print "Car Rental Expense:", FormatCurrency(carRentaltotal)
picResults.Print "*************************************************************************************"
picResults.Print "Trip Total Expense:", FormatCurrency(Finaltotal)

cmdend.Visible = True

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
