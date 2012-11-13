VERSION 5.00
Begin VB.Form frmSeventh 
   Caption         =   "Second Run"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   9615
      Left            =   -480
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   9555
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   -840
      Width           =   11775
      Begin VB.PictureBox picResults5 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "JazzTextExtended"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1080
         ScaleHeight     =   2595
         ScaleWidth      =   3435
         TabIndex        =   16
         Top             =   6480
         Width           =   3495
      End
      Begin VB.CommandButton cmdTailgrab 
         BackColor       =   &H00C0FFC0&
         Caption         =   "360 Tailgrab"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdMethod 
         BackColor       =   &H000000FF&
         Caption         =   "1080 Method"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton cmdHaakon 
         BackColor       =   &H000080FF&
         Caption         =   "Haakon Flip"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdRoastbeef 
         BackColor       =   &H0080FFFF&
         Caption         =   "540 Roastbeef"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdSato 
         BackColor       =   &H000000FF&
         Caption         =   "Sato Flip"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton cmdMcTwist 
         BackColor       =   &H0080FFFF&
         Caption         =   "McTwist"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdEggplant 
         BackColor       =   &H0080FFFF&
         Caption         =   "Eggplant"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdRodeo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rodeo"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdMelon 
         BackColor       =   &H000080FF&
         Caption         =   "720 Melon"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdInvert 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Invert"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCrippler 
         BackColor       =   &H000000FF&
         Caption         =   "Crippler"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton cmdEggflip 
         BackColor       =   &H000080FF&
         Caption         =   "Eggflip"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Next-->"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8160
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Invert Tricks"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Flip Tricks"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5160
         TabIndex        =   18
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Spin/ Grabs"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Tricks:"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   840
         TabIndex        =   2
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Second of 3 Runs"
         BeginProperty Font 
            Name            =   "Chiller"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmSeventh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer, Pos As Integer
Dim Found As Boolean

Dim tailgrab As Single, roastbeef As Single, melon As Single, method As Single
Dim rodeo As Single, mctwist As Single, haakon As Single, sato As Single
Dim invert As Single, eggplant As Single, eggflip As Single, crippler As Single

Private Sub cmdNext_Click()

frmSeventh.Hide
frmEight.Show

End Sub

Private Sub cmdCrippler_Click()
crippler = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Crippler"
    TrickPoints2(CTR) = crippler
Else
    TrickName2(CTR) = "Crippler"
    TrickPoints2(CTR) = crippler
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If
    
If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub
Private Sub cmdEggflip_Click()
eggflip = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Eggflip"
    TrickPoints2(CTR) = eggflip
Else
    TrickName2(CTR) = "Eggflip"
    TrickPoints2(CTR) = eggflip
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If

End Sub
Private Sub cmdEggplant_Click()
eggplant = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Eggplant"
    TrickPoints2(CTR) = eggplant
Else
    TrickName2(CTR) = "Eggplant"
    TrickPoints2(CTR) = eggplant
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub
Private Sub cmdInvert_Click()
invert = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Invert"
    TrickPoints2(CTR) = invert
Else
    TrickName2(CTR) = "Invert"
    TrickPoints2(CTR) = invert
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub
Private Sub cmdHaakon_Click()
haakon = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Haakon Flip"
    TrickPoints2(CTR) = haakon
Else
    TrickName2(CTR) = "Haakon Flip"
    TrickPoints2(CTR) = haakon
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    
End Sub
Private Sub cmdMcTwist_Click()
mctwist = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "McTwist"
    TrickPoints2(CTR) = mctwist
Else
    TrickName2(CTR) = "McTwist"
    TrickPoints2(CTR) = mctwist
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub
Private Sub cmdMethod_Click()
method = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "1080 Method"
    TrickPoints2(CTR) = method
Else
    TrickName2(CTR) = "1080 Method"
    TrickPoints2(CTR) = method
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub

Private Sub cmdMelon_Click()
melon = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "720 Melon"
    TrickPoints2(CTR) = melon
Else
    TrickName2(CTR) = "720 Melon"
    TrickPoints2(CTR) = melon
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub

Private Sub cmdRodeo_Click()
rodeo = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Rodeo"
    TrickPoints2(CTR) = rodeo
Else
    TrickName2(CTR) = "Rodeo"
    TrickPoints2(CTR) = rodeo
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    
    
End Sub

Private Sub cmdSato_Click()
sato = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "Sato Flip"
    TrickPoints2(CTR) = sato
Else
    TrickName2(CTR) = "Sato Flip"
    TrickPoints2(CTR) = sato
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    
    
End Sub

Private Sub cmdTailgrab_Click()
tailgrab = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "360 Tailgrab"
    TrickPoints2(CTR) = tailgrab
Else
    TrickName2(CTR) = "360 Tailgrab"
    TrickPoints2(CTR) = tailgrab
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    
    
End Sub
Private Sub cmdRoastbeef_Click()
roastbeef = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName2(CTR) = "540 Roastbeef"
    TrickPoints2(CTR) = roastbeef
Else
    TrickName2(CTR) = "540 Roastbeef"
    TrickPoints2(CTR) = roastbeef
    'disable all buttons
    cmdTailgrab.Enabled = False
    cmdRoastbeef.Enabled = False
    cmdMelon.Enabled = False
    cmdMethod.Enabled = False
    cmdRodeo.Enabled = False
    cmdMcTwist.Enabled = False
    cmdHaakon.Enabled = False
    cmdSato.Enabled = False
    cmdInvert.Enabled = False
    cmdEggplant.Enabled = False
    cmdEggflip.Enabled = False
    cmdCrippler.Enabled = False
End If

If Found = False Then
    picResults5.Print "Trick"; Tab(20); "Points"
    picResults5.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults5.Print TrickName2(CTR); Tab(20); TrickPoints2(CTR)
End If
    

End Sub

