VERSION 5.00
Begin VB.Form frmSixth 
   Caption         =   "First Run"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   11640
      Left            =   -3600
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   11580
      ScaleWidth      =   15420
      TabIndex        =   0
      Top             =   -2880
      Width           =   15480
      Begin VB.PictureBox picResults4 
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
         Height          =   2535
         Left            =   3840
         ScaleHeight     =   2475
         ScaleWidth      =   3435
         TabIndex        =   16
         Top             =   8520
         Width           =   3495
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6600
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7200
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6600
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
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
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6000
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6000
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7200
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6000
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
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6600
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7200
         Width           =   1935
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9840
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
         Left            =   11400
         TabIndex        =   19
         Top             =   4800
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
         Left            =   7680
         TabIndex        =   18
         Top             =   4800
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
         Left            =   4080
         TabIndex        =   17
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Tricks: (Maximum of 4 Tricks)"
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
         Height          =   975
         Left            =   3840
         TabIndex        =   2
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "First of 3 Runs"
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
         Height          =   975
         Left            =   3840
         TabIndex        =   1
         Top             =   3000
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmSixth"
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

frmSixth.Hide
frmSeventh.Show

End Sub

Private Sub cmdCrippler_Click()
crippler = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Crippler"
    TrickPoints1(CTR) = crippler
Else
    TrickName1(CTR) = "Crippler"
    TrickPoints1(CTR) = crippler
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
Private Sub cmdEggflip_Click()
eggflip = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Eggflip"
    TrickPoints1(CTR) = eggflip
Else
    TrickName1(CTR) = "Eggflip"
    TrickPoints1(CTR) = eggflip
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If

End Sub
Private Sub cmdEggplant_Click()
eggplant = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Eggplant"
    TrickPoints1(CTR) = eggplant
Else
    TrickName1(CTR) = "Eggplant"
    TrickPoints1(CTR) = eggplant
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
Private Sub cmdInvert_Click()
invert = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Invert"
    TrickPoints1(CTR) = invert
Else
    TrickName1(CTR) = "Invert"
    TrickPoints1(CTR) = invert
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
Private Sub cmdHaakon_Click()
haakon = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Haakon Flip"
    TrickPoints1(CTR) = haakon
Else
    TrickName1(CTR) = "Haakon Flip"
    TrickPoints1(CTR) = haakon
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
Private Sub cmdMcTwist_Click()
mctwist = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "McTwist"
    TrickPoints1(CTR) = mctwist
Else
    TrickName1(CTR) = "McTwist"
    TrickPoints1(CTR) = mctwist
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
Private Sub cmdMethod_Click()
method = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "1080 Method"
    TrickPoints1(CTR) = method
Else
    TrickName1(CTR) = "1080 Method"
    TrickPoints1(CTR) = method
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub

Private Sub cmdMelon_Click()
melon = 15
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "720 Melon"
    TrickPoints1(CTR) = melon
Else
    TrickName1(CTR) = "720 Melon"
    TrickPoints1(CTR) = melon
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub

Private Sub cmdRodeo_Click()
rodeo = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Rodeo"
    TrickPoints1(CTR) = rodeo
Else
    TrickName1(CTR) = "Rodeo"
    TrickPoints1(CTR) = rodeo
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    
    
End Sub

Private Sub cmdSato_Click()
sato = 20
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "Sato Flip"
    TrickPoints1(CTR) = sato
Else
    TrickName1(CTR) = "Sato Flip"
    TrickPoints1(CTR) = sato
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    
    
End Sub

Private Sub cmdTailgrab_Click()
tailgrab = 5
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "360 Tailgrab"
    TrickPoints1(CTR) = tailgrab
Else
    TrickName1(CTR) = "360 Tailgrab"
    TrickPoints1(CTR) = tailgrab
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    
    
End Sub
Private Sub cmdRoastbeef_Click()
roastbeef = 10
CTR = CTR + 1
If CTR < 4 Then
    TrickName1(CTR) = "540 Roastbeef"
    TrickPoints1(CTR) = roastbeef
Else
    TrickName1(CTR) = "540 Roastbeef"
    TrickPoints1(CTR) = roastbeef
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
    picResults4.Print "Trick"; Tab(20); "Points"
    picResults4.Print "*********************************************"
    Found = True
End If

If Found = True Then
    'print out both arrays
    picResults4.Print TrickName1(CTR); Tab(20); TrickPoints1(CTR)
End If
    

End Sub
