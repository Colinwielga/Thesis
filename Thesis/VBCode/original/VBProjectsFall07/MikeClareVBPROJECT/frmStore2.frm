VERSION 5.00
Begin VB.Form frmStore2 
   BackColor       =   &H0000C000&
   Caption         =   "Old Man's Store"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoney 
      Caption         =   "View my stats!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton cmdKnife 
      Height          =   1335
      Left            =   2880
      Picture         =   "frmStore2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdShield 
      Height          =   1455
      Left            =   5280
      Picture         =   "frmStore2.frx":0772
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdTaser 
      Height          =   1815
      Left            =   7800
      Picture         =   "frmStore2.frx":60C4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave Store"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   " The Old Man's Store"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2640
      TabIndex        =   14
      Top             =   360
      Width           =   7935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Pick out an item to buy, click 'View my Stats' to see your Money, H.P., and Attack."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   13
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ITEMS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KNIFE: $600"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack + 50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TASER: $1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack + 75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SHIELD: $400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "H.P. + 75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "frmStore2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdKnife_Click()    'same as store1, but prices rose.  knife is $600 now.
    
    
    If Money > 600 Then
        Money = Money - 600
        Attack = Attack + 50
        MsgBox ("You bought a knife.  Now your attack is at: " & Attack & ".  You used $600 and have " & FormatCurrency(Money) & " left over."), , ("Knife")
    Else
        MsgBox ("You don't have enough cash"), , ("Error")
    End If
    
   cmdKnife.Enabled = False
   
End Sub

Private Sub cmdLeave_Click() 'leave store
    frmStore2.Hide
    frmStreet2.Show
    
End Sub

Private Sub cmdMoney_Click()    'see your current stats
    picResults.Cls

    picResults.Print "You have: "; FormatCurrency(Money)
    picResults.Print "You have: "; HP; " H.P."
    picResults.Print "You have: "; Attack; " attack points"
End Sub

Private Sub cmdShield_Click() 'shield is $400 now
    If Money > 400 Then
        Money = Money - 400
        HP = HP + 75
        MsgBox ("You bought a shield.  Now your H.P. is at: " & HP & ".  You used $400 and have " & FormatCurrency(Money) & " left over."), , ("Shield")
    Else
        MsgBox ("You don't have enough cash"), , ("Error")
    End If
    
    cmdShield.Enabled = False
    
End Sub

Private Sub cmdTaser_Click()    'taser is $1000 now.
    If Money > 1000 Then
        Money = Money - 1000
        Attack = Attack + 75
        MsgBox ("You bought a taser.  Now your attack is at: " & Attack & ".  You used $1000 and have " & FormatCurrency(Money) & " left over."), , ("Taser")
    Else
        MsgBox ("You don't have enough cash"), , ("Error")
    End If
    
   cmdTaser.Enabled = False
End Sub


