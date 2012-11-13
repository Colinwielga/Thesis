VERSION 5.00
Begin VB.Form Checkout 
   BackColor       =   &H00000000&
   Caption         =   "Checkout"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "Purchase Another Ski Package"
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Display Total"
      Height          =   855
      Left            =   3960
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox pbxResults 
      Height          =   4935
      Left            =   480
      ScaleHeight     =   4875
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Checkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: Checkout (Checkout.frm)
'Author David Meacham
'Date Written: Wednesday, October 23
'Purpose of form: Calculates the final total of the ski's, boot's
                    'and bindings together.  Then calculates tax
                    'and prints out a final total.  Also allows
                    'the user to quit the program.

Option Explicit

Private Sub cmdNew_Click()
'Allows the user to start over and purchase a new ski package
Checkout.Hide
LevelSelect.Show
End Sub

Private Sub cmdQuit_Click()
'ends program
End
End Sub

Private Sub cmdTotal_Click()
'Adds the total from the ski's, boots, and bindings
pbxResults.Cls
Dim tax As Single
Dim ftotal As Single
tax = sum * 0.065                   'calculates tax on the subtotal
ftotal = sum + tax
pbxResults.Print "SubTotal", FormatCurrency(sum)
pbxResults.Print "Tax", FormatCurrency(tax)
pbxResults.Print "************************"
pbxResults.Print "Total", FormatCurrency(ftotal)
pbxResults.Print
pbxResults.Print "Thank's for shopping Dave's Ski Shop!"
End Sub

