VERSION 5.00
Begin VB.Form dsfds
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMilkshakes
      BackColor       =   &H000080FF&
      Caption         =   "Milkshakes"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdOyster
      BackColor       =   &H000080FF&
      Caption         =   "Oysters"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdPancakes
      BackColor       =   &H000080FF&
      Caption         =   "Pancakes"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdPizza
      BackColor       =   &H000080FF&
      Caption         =   "Pizza"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeli
      BackColor       =   &H000080FF&
      Caption         =   "Deli"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdBar
      BackColor       =   &H000080FF&
      Caption         =   "Bar"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOmelet
      BackColor       =   &H000080FF&
      Caption         =   "Omelet"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBack
      BackColor       =   &H0080FF80&
      Caption         =   "<===Go Back"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSteak
      BackColor       =   &H000080FF&
      Caption         =   "Steak"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdBurger
      BackColor       =   &H000080FF&
      Caption         =   "Burger"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblChoose
      BackColor       =   &H000000FF&
      Caption         =   "Choose your meal by clicking on a button under the city lables or to exit click on quit."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2760
      TabIndex        =   20
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label lblAmarillo
      BackColor       =   &H00FFFFC0&
      Caption         =   "9. Amarillo"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   17
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblStLouis
      BackColor       =   &H00FFFFC0&
      Caption         =   "8. St. Louis"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   16
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblNewOrleans
      BackColor       =   &H00FFFFC0&
      Caption         =   "7. New             Orleans"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblHawaii
      BackColor       =   &H00FFFFC0&
      Caption         =   "6. Hawaii"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblAtlanta
      BackColor       =   &H00FFFFC0&
      Caption         =   "5. Atlanta"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblBoston
      BackColor       =   &H00FFFFC0&
      Caption         =   "4. Boston"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   12
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblDetroit
      BackColor       =   &H00FFFFC0&
      Caption         =   "3. Detroit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label lblMinneapolis
      BackColor       =   &H00FFFFC0&
      Caption         =   "2. Minneapolis"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblSeattle
      BackColor       =   &H00FFFFC0&
      Caption         =   "1. Seattle"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl9
      BackColor       =   &H000000FF&
      Caption         =   "9."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label lbl8
      BackColor       =   &H000000FF&
      Caption         =   "8."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lbl7
      BackColor       =   &H000000FF&
      Caption         =   "7."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lbl6
      BackColor       =   &H000000FF&
      Caption         =   "6."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   6000
      Width           =   255
   End
   Begin VB.Label lbl5
      BackColor       =   &H000000FF&
      Caption         =   "5."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lbl4
      BackColor       =   &H000000FF&
      Caption         =   "4."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lbl3
      BackColor       =   &H000000FF&
      Caption         =   "3."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lbl2
      BackColor       =   &H000000FF&
      Caption         =   "2."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbl1
      BackColor       =   &H000000FF&
      Caption         =   "1."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgmap
      Height          =   6795
      Left            =   2040
      Picture         =   "dsfds.frx":0000
      Top             =   240
      Width           =   11250
   End
End
Attribute VB_Name = "dsfds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Man vs.Food
'dsfds
'Ty Nimens and Josh Seaburg
'February 2010
'This form has all of our buttons that when clicked it leads to another form so the user can learn more about a Man vs. Food episode

Private Sub asdf_Click()
    wefew.Show
    dsfds.Hide
End Sub

Private Sub sdfg_Click()
    dsfds.Hide
    fwefwef.Show
End Sub

Private Sub dfgh_Click()
    frmDeli.Show
    dsfds.Hide
End Sub

Private Sub fghj_Click()
    dsfds.Hide
    bvcv.Show
End Sub

Private Sub ghjk_Click()
   ' the codes shown below hide the map when a button is clicked and show the page indicated on the button the users clicks on
    dsfds.Hide
    rtht.Show
End Sub

Private Sub hjkl_Click()
    frmOmelet.Show
    dsfds.Hide
End Sub

Private Sub pou_Click()
    frmOyster.Show
    dsfds.Hide
End Sub

Private Sub oiuy_Click()
    hhhhh.Show
    dsfds.Hide
End Sub

Private Sub cmdPizza_Click()
    rthnteh.Show
    dsfds.Hide
End Sub

Private Sub iuyt_Click()
    End
End Sub

Public Sub uytr_Click()
    aqaqaq.Show
    dsfds.Hide

End Sub
