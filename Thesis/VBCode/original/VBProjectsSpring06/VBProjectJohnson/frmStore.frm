VERSION 5.00
Begin VB.Form frmStore 
   BackColor       =   &H00FF0000&
   Caption         =   "School Store"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   7080
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   18
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H000000FF&
      Caption         =   "Submit Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   1455
   End
   Begin VB.PictureBox picChair 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   4800
      Picture         =   "frmStore.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   1575
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox picFridge 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   4800
      Picture         =   "frmStore.frx":08BD
      ScaleHeight     =   1695
      ScaleWidth      =   1815
      TabIndex        =   13
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox picLamp 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      Height          =   1575
      Left            =   480
      Picture         =   "frmStore.frx":143D
      ScaleHeight     =   1575
      ScaleWidth      =   1695
      TabIndex        =   12
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox picFuton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      Height          =   1455
      Left            =   480
      Picture         =   "frmStore.frx":1FFE
      ScaleHeight     =   1455
      ScaleWidth      =   1815
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.PictureBox picLofts 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   480
      Picture         =   "frmStore.frx":2B27
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtChair 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   7440
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtFridge 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   7320
      TabIndex        =   8
      Text            =   "0"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtLamp 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Text            =   "0"
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtFuton 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Text            =   "0"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtLofts 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00FF0000&
      Caption         =   "Project Created By:                 Kyle Johnson"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8280
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FF0000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblDesk 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Quantity Of  Desk Chairs Needed ( $ 49.99 each)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblFridge 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter The Quantity Of Refridgerators Needed ($ 239.00 each)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblLamp 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Quantity Of Lamps Needed ( $ 8.67 each)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblFuton 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Quantity Of Futons Needed ($ 124.99 each)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblLoft 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter Quantity Of Lofts Needed ($87.99 each)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. Johns Housing Project
' Store Form
' Written By Kyle Johnson
' 3/22/06
' This form allows the user to purchase items necessary for thier house
' it calculates the total of thier purchase, and displays an appropriate message box




Option Explicit
    'dim global variables
    Dim Futon, Lofts, Chair, Lamp, Fridge As Integer
    Dim Total As Single


Private Sub cmdBack_Click()
    'navigates from the store to the options page
    frmStore.Visible = False
    frmOptions.Visible = True
    
End Sub

Private Sub cmdSubmit_Click()
    'sets the variables equal to the appropriate input box in order to calculate the total
    Futon = txtFuton.Text
    Lofts = txtLofts.Text
    Chair = txtChair.Text
    Lamp = txtLamp.Text
    Fridge = txtFridge.Text
    'calculate the total of the order made
    Total = (124.99 * Futon) + (87.99 * Lofts) + (49.99 * Chair) + (8.67 * Lamp) + (239 * Fridge)
    'print the total in the picture box
    picResults.Print FormatCurrency(Total)
    
    'give an appropriate message box based on the size of the order
    Select Case Total
        Case 0 To 10
            MsgBox namesArray(K) & " You are A Terrible Customer", , "Dont Come Back"
        Case 10 To 50
            MsgBox namesArray(K) & " Thanks For Nothing", , "The Store Is Closed"
        Case 50 To 100
            MsgBox namesArray(K) & " You Are A Fine Customer", , " Come Back Anytime"
        Case Is > 100
            MsgBox namesArray(K) & " Your Putting My Kids Through College", , "Thank You"
    End Select
End Sub


