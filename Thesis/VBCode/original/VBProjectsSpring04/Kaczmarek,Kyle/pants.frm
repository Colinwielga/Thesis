VERSION 5.00
Begin VB.Form frmpants 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Pants"
   ClientHeight    =   8580
   ClientLeft      =   1320
   ClientTop       =   1515
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdcleats 
      BackColor       =   &H00FF8080&
      Caption         =   "Order Baseball Cleats"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdclearpants 
      BackColor       =   &H00FF8080&
      Caption         =   "Clear Calculations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdcalcpants 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   3135
   End
   Begin VB.PictureBox picpants 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   8040
      ScaleHeight     =   3675
      ScaleWidth      =   3675
      TabIndex        =   10
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox txtpants 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   4800
      Width           =   2535
   End
   Begin VB.OptionButton optwhitestripe 
      BackColor       =   &H00C0E0FF&
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
      Left            =   10800
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton optpinstripe 
      BackColor       =   &H00C0E0FF&
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
      Left            =   7800
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton optgreystripe 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton optgrey 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picwhitestripe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   9720
      Picture         =   "pants.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.PictureBox picgreystripe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3600
      Picture         =   "pants.frx":1C7E2
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox picgrey 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   480
      Picture         =   "pants.frx":39CE4
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox picpinstripe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6720
      Picture         =   "pants.frx":571E6
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblwhite 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "$54.95"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   21
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblpinstripes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "$49.99"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblgreystriped 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "$54.95"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblgrey 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "$29.99"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "White Striped"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   9840
      TabIndex        =   17
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pinstripes"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grey Striped"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grey Pants"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "How Many Pairs of Pants?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmpants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BaseballUniforms (BaseballUniforms.vbp)
'Form Name : frmpants (pants.frm)
'Author: Kyle Kaczmarek
'Date Written: March 15, 2004
'Purpose of Form:'



Private Sub cmdcalcpants_Click()

pants = txtpants.Text 'enter how many pairs of pants you want


If optgrey = True Then 'chosen the grey option
        pantsprice = 29.99 'grey price
    ElseIf optgreystripe = True Then 'chosen the grey stripe option
        pantsprice = 54.95 'grey option price
    ElseIf optpinstripe = True Then 'chosen the pin stripe option
        pantsprice = 49.99 'pin stripe price
    ElseIf optwhitestripe = True Then 'chosen the white stripe option
        pantsprice = 54.95 'white stripe price

End If

pantstotal = pants * pantsprice 'multiplies the number of pants by the price
    
picpants.Cls 'clears the picture box
picpants.Print "Pairs of Pants", "Cost" 'prints out the titles
picpants.Print "********************", "*****" 'prints out the stars
picpants.Print Tab(8); pants, , FormatCurrency(pantstotal, 2) 'prints out the total cost for the pants
End Sub

Private Sub cmdclearpants_Click()

picpants.Cls 'clear the picture box
txtpants = "" 'puts nothing into the inout box

End Sub

Private Sub cmdcleats_Click()
frmjerseys.Hide 'closes the jerseys form
frmhats.Hide 'closes the hat form
frmorder.Hide 'closes the order form
frmpants.Hide 'closes the pants form
frmcleats.Show 'shows the cleats form
frmfinal.Hide 'closes the final form
End Sub

