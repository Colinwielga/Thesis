VERSION 5.00
Begin VB.Form frmEntrance 
   Caption         =   "Target"
   ClientHeight    =   8940
   ClientLeft      =   3495
   ClientTop       =   3180
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   Picture         =   "frmEntrance.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   10875
   Begin VB.CommandButton cmdBeginShopping 
      Caption         =   "Begin Shopping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   7680
      Width           =   4215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   0
      Top             =   7680
      Width           =   4215
   End
End
Attribute VB_Name = "frmEntrance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBeginShopping_Click()
Dim X As Integer
'This subroutine hides the entrance form and shows the department selection form
frmEntrance.Hide
frmDepartments.Show
'it also opens each department's cart and erases the file by printing blanks in locations 0 to 100.
'excess blanks do not affect the program as the do not match with any inventory number in the 'master list' during checkout.

'This reenables the display button in the checkout form if this is a second shopping trip.
frmCheckout.cmdDisplay.Enabled = True

Open App.Path & "\ToysCart.txt" For Output As #1

For X = 1 To 100
    Print #1, ""
Next X
Close #1

Open App.Path & "\KitchenCart.txt" For Output As #2

For X = 1 To 100
    Print #2, ""
Next X
Close #2

Open App.Path & "\FurnitureCart.txt" For Output As #3

For X = 1 To 100
    Print #3, ""
Next X
Close #3

Open App.Path & "\ElectronicsCart.txt" For Output As #4

For X = 1 To 100
    Print #4, ""
Next X
Close #4

Open App.Path & "\ClothingCart.txt" For Output As #5

For X = 1 To 100
    Print #5, ""
Next X
Close #5

End Sub

Private Sub cmdQuit_Click()
'This subroutine quits out of the program and also sends the user a msgbox.
MsgBox "Have a great day!", , "Target"
End

End Sub

Private Sub Form_Load()
'This subroutine greets the user with a msgbox and the Entrance form.
frmEntrance.Show
MsgBox "Hello! Welcome to Target, thank you for choosing us for your shopping needs.", , "Target"

End Sub
