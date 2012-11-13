VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000001&
   Caption         =   "What's on the Menu??"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   8040
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   3840
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   4680
      Picture         =   "frmMenu.frx":2102
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   3840
      Width           =   2415
   End
   Begin VB.PictureBox picResultss 
      BackColor       =   &H00008000&
      FillColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   4680
      ScaleHeight     =   1035
      ScaleWidth      =   5715
      TabIndex        =   9
      Top             =   2640
      Width           =   5775
   End
   Begin VB.CommandButton cmdDessert 
      BackColor       =   &H8000000B&
      Caption         =   "Click here to see the dessert menu!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   5760
      Width           =   3735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "See the total with tax and tip!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C000&
      FillColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox txtOrder 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H8000000B&
      Caption         =   "Click here to see the menu!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "Please enter your order (exactly as written):"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   " What's on the Menu??"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   $"frmMenu.frx":5F36
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10095
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'What's on the Menu?
'frmMenu
'Michael Murakami and Kevin Schnese
'25 October 2006
'This form allows the user to search the menu.  The user can look over the menu select a meal. The program will then output what the user selected and what the total is with tax and tip included. This form also allows the user to go to the dessert form.

Private Sub cmdDessert_Click()
'This will allow the user to switch between the regular menu and the dessert menu.
    frmMenu.Visible = False
    frmDessert.Visible = True
End Sub
Private Sub cmdExit_Click()
'A message box will pop up after clicking the exit button and and thank the user for using our menu.
    MsgBox "Thank you!", , "THANK YOU!!!"
    End
End Sub
Private Sub cmdMenu_Click()
'This set of code, once the button is pushed, will pull a set of data from a notebook and and print it in the picture box.
    counter = 0
    picResults.Cls
    picResults.Print "The items on the menu are:"
    picResults.Print "************************************"
    Open App.Path & "\menu.txt" For Input As #1
        Do Until EOF(1)
        Input #1, fooditem, price
        counter = counter + 1
        fooditems(counter) = fooditem
        foodprices(counter) = price
        picResults.Print counter; ".) "; fooditem; " for "; FormatCurrency(price)
        picResults.Print " "
    Loop
    Close #1
End Sub
Private Sub cmdSearch_Click()
'This set of code will pull the meal that was entered in the text box and and print out what they selected and then do the equation for the tax and tip and then print it out in the picture box.
    counter = 0
    newcounter = 0
    tip = 0.15
    tax = 0.065
    Found = False
    order = txtOrder.Text
    picResultss.Cls
    Open App.Path & "\menu.txt" For Input As #1
        Do Until Found = True Or EOF(1)
            Input #1, fooditem, price
                counter = counter + 1
                fooditems(counter) = fooditem
                foodprices(counter) = price
                If order = fooditem Then
                    Found = True
                    desired = fooditem
                    taxed = price * 0.065
                    tipped = price * 0.15
                    total = price + taxed + tipped
                    picResultss.Print "You have chosen "; desired; " at a price of "; FormatCurrency(price)
                    picResultss.Print " "
                    picResultss.Print "The total after a tax of 6.5% and a tip of 15% is: "; FormatCurrency(total)
                Else
                    Found = False
                End If
        Loop
        Close #1
        If (Not Found) Then
            MsgBox "Please check your spelling and try again!", , "Try Again"
        End If
End Sub

