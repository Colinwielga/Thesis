VERSION 5.00
Begin VB.Form frmMatchPrice 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Kim Nguyen"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResultMatch 
      BackColor       =   &H00FFC0FF&
      Height          =   5295
      Left            =   840
      ScaleHeight     =   5235
      ScaleWidth      =   8955
      TabIndex        =   5
      Top             =   840
      Width           =   9015
   End
   Begin VB.PictureBox picResultMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   6480
      Width           =   9015
   End
   Begin VB.CommandButton cmdLoanPayment 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Loan Payment Calculator"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Return To Homepage"
      Height          =   735
      Left            =   5640
      TabIndex        =   2
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdFindMatch 
      Caption         =   "Find A Car That Match My Price"
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   7200
      Width           =   2055
   End
End
Attribute VB_Name = "frmMatchPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Online Car Dealer
'Form Name : frmMatchPrice (FrmMatchPrice.frm)
'Author: Kim Nguyen
'Date Written: October 29, 2003
'Purpose of Form: This form will let the user search for a car
                    'that match their price using the input box
                'when the user enter the price that fit in a certain case
                'a picture of that car will print out in the picture box

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.Option Explicit
Option Explicit
'Declares all the varialbes
Dim Price As Single
Dim I As Integer
Dim Picure As String
Dim CarPic(1 To 12) As String


'Ask the user for the input than select the case that fit the price and print the
'picture of that car in the price range in the picture box


Private Sub cmdFindMatch_Click()
'declare the variables
Dim Price As Double
'clear the picture box so that the user can do it again
picResultMatch.Cls
picResultMessage.Cls


Price = InputBox("Enter Price Range from 15,000 to 80,000")


'cases that the user have to use to see if their price fit in an apporiate range
'if it does, a car picture of that range will be printed out other wise
'a message box will pop up and let the user know that they must enter the price with in the
'price range
Select Case Price
        Case Is >= 80000
            picResultMatch.Picture = LoadPicture(Path & "Car80000.jpg")
        Case 70000 To 79999
            picResultMatch.Picture = LoadPicture(Path & "Car70000.jpg")
        Case 65000 To 69999
            picResultMatch.Picture = LoadPicture(Path & "Car65000.jpg")
        Case 60000 To 64999
            picResultMatch.Picture = LoadPicture(Path & "Car60000.jpg")
        Case 55000 To 59999
            picResultMatch.Picture = LoadPicture(Path & "Car55000.jpg")
        Case 50000 To 54999
            picResultMatch.Picture = LoadPicture(Path & "Car50000.jpg")
        Case 45000 To 49999
            picResultMatch.Picture = LoadPicture(Path & "Car45000.jpg")
        Case 40000 To 49999
            picResultMatch.Picture = LoadPicture(Path & "Car40000.jpg")
        Case 35000 To 39999
            picResultMatch.Picture = LoadPicture(Path & "Car35000.jpg")
        Case 30000 To 34999
            picResultMatch.Picture = LoadPicture(Path & "Car30000.jpg")
        Case 25000 To 29999
            picResultMatch.Picture = LoadPicture(Path & "Car25000.jpg")
        Case 20000 To 24999
            picResultMatch.Picture = LoadPicture(Path & "Car20000.jpg")
        Case 15000 To 19999
            picResultMatch.Picture = LoadPicture(Path & "Car15000.jpg")
        Case Is < 15000
            picResultMessage.Print "With"; " "; FormatCurrency(Price, 2); " "; "You Can't Buy Our Cars!"
            MsgBox "Sorry You Must Enter the Number With The Range", , "Error"
            
    End Select
        picResultMessage.Print "With"; " "; FormatCurrency(Price, 2); " "; "You Can Buy This Car!"
    

            
   

        
Close #3
End Sub
'Show the Info page  and hide the MatchPrice page
Private Sub cmdHomepage_Click()
frmInfo.Show
frmMatchPrice.Hide
End Sub

'show the loan payment calculator so that the user can use
Private Sub cmdLoanPayment_Click()
frmLoan.Show
End Sub
'End the program now
Private Sub cmdQuit_Click()
End
End Sub

Private Sub picResultMatch_Click()

End Sub
