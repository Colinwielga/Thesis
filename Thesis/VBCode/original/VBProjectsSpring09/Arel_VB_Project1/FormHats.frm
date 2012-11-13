VERSION 5.00
Begin VB.Form FrmHats 
   BackColor       =   &H00400000&
   Caption         =   "Form2"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Order"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate Total"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton CmdOrder9 
      BackColor       =   &H000000FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder8 
      BackColor       =   &H000000C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder7 
      BackColor       =   &H000000FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder6 
      BackColor       =   &H000000C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder5 
      BackColor       =   &H00000080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder4 
      BackColor       =   &H000000C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox picResults 
      Height          =   4215
      Left            =   5640
      ScaleHeight     =   4155
      ScaleWidth      =   3555
      TabIndex        =   19
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton CmdOrder3 
      BackColor       =   &H000000FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton CmdOrder2 
      BackColor       =   &H000000C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox PicPrice 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   16
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton CmdOrder1 
      BackColor       =   &H000000FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   5655
      Left            =   3600
      ScaleHeight     =   5595
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox picHat 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   720
      ScaleHeight     =   2235
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton CmdHat9 
      Caption         =   "9"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton CmdHat8 
      Caption         =   "8"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton CmdHat7 
      Caption         =   "7"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton CmdHat6 
      Caption         =   "6"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdHat5 
      Caption         =   "5"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdHat4 
      Caption         =   "4"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton CmdHat3 
      Caption         =   "3"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton CmdHat2 
      Caption         =   "2"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton CmdHat1 
      Caption         =   "1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000D&
      Height          =   2415
      Left            =   600
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project By: Steph Arel"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order Here!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER HERE! "
      Height          =   735
      Left            =   5280
      TabIndex        =   15
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Twins Hat Central!"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here To See Each Hat!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "FrmHats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Twins Fan
'Form Name: FrmHats
'Project By: Stephanie Arel
'Date Written: 3/12/2009
'This Form allows the user to view each hat and then purchase the ones they like.
Option Explicit

Dim SubTotal As Integer
Dim Total As Integer
Dim Tax As Integer




Private Sub CmdHat1_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 1a.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $35"



End Sub

Private Sub CmdHat2_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 2.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $35"
End Sub

Private Sub CmdHat3_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 3.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $35"
End Sub

Private Sub CmdHat4_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 4.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $35"
End Sub

Private Sub CmdHat5_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 5.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $30"
End Sub

Private Sub CmdHat6_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 6.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $30"
End Sub

Private Sub CmdHat7_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 7.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $25"
End Sub

Private Sub CmdHat8_Click()

'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 8.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $25"
End Sub

Private Sub CmdHat9_Click()
'Clears the hat picture box as well as the price picture box
picHat.Cls
PicPrice.Cls
'Loads the selected picture and price of the hat.
picHat.Picture = LoadPicture(App.Path & "/Hat 9.jpg")
PicPrice.ForeColor = vbRed
PicPrice.Print "Price: $25"
End Sub




Private Sub CmdOrder1_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 1         $35"

SubTotal = 35 + SubTotal

End Sub

Private Sub CmdOrder2_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 2         $35"

SubTotal = 35 + SubTotal
End Sub

Private Sub CmdOrder3_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 3         $35"

SubTotal = 35 + SubTotal
End Sub

Private Sub CmdOrder4_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 4         $35"

SubTotal = 35 + SubTotal
End Sub

Private Sub CmdOrder5_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 5         $30"

SubTotal = 30 + SubTotal
End Sub

Private Sub CmdOrder6_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 6         $30"

SubTotal = 30 + SubTotal
End Sub

Private Sub CmdOrder7_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 7         $25"

SubTotal = 25 + SubTotal
End Sub

Private Sub CmdOrder8_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 8         $25"

SubTotal = 25 + SubTotal
End Sub

Private Sub CmdOrder9_Click()
'Prints the selection and adds the price to the subtotal.
picResults.Print "Hat 9         $25"

SubTotal = 25 + SubTotal
End Sub

Private Sub CmdQuit_Click()
'Ends the program completely
End

End Sub

Private Sub Command1_Click()
'Prints the user's purchases with tax.
Tax = SubTotal * 0.07
Total = SubTotal + Tax

picResults.Print
picResults.Print "**********************************"
picResults.Print
picResults.Print "SubTotal:          "; FormatCurrency(SubTotal, 2)
picResults.Print "Tax:                  "; FormatCurrency(Tax, 2)
picResults.Print "**********************************"
picResults.Print "Total:                "; FormatCurrency(Total, 2)
End Sub

Private Sub Command2_Click()
'Returns to main menu
FrmHats.Hide
FrmMain.Show

End Sub

Private Sub Command3_Click()
'Clears the user's selections and starts over.
picResults.Cls
SubTotal = 0
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
