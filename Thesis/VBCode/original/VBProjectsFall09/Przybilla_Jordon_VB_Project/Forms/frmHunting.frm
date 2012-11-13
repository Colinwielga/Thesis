VERSION 5.00
Begin VB.Form frmHunting 
   Caption         =   "Hunting"
   ClientHeight    =   11130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   Picture         =   "frmHunting.frx":0000
   ScaleHeight     =   11130
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrip 
      Caption         =   "Plan A Hunting Trip in Minnesota"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8280
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdTIP 
      Caption         =   "#1 Rule for hunting of any kind."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   4935
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4560
      TabIndex        =   3
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10320
      Width           =   975
   End
   Begin VB.CommandButton cmdRifle 
      Caption         =   "Rifle  Hunting"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblHunting 
      BackColor       =   &H8000000C&
      Caption         =   "  Push a button to find out more about that type of hunting."
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmHunting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MN Deer'
'Form Name: Hunting'
'Authors: Jordon Przybilla'
'Date Written: October 8, 2009
'this form will be like a sub-home page, from here the user will be able to learn about rifle hunting and also be able
'to set up a hunting trip at Whitetail X-treme Hunts


Option Explicit

Private Sub cmdBow_Click() 'takes the user to the bow & muzzleloader hunting form

frmHunting.Hide
frmBowMuzzle.Show

End Sub

Private Sub cmdHome_Click() ' takes the user back to the home page


frmHunting.Hide
frmHome.Show


End Sub

Private Sub cmdMuzzle_Click() 'takes the user to the muzzleloader hunting form

frmHunting.Hide
frmMuzzleloading.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRifle_Click() 'takes the user to the rifle and shotgun hunting form, and loads arrays



Open App.Path & "\Data\rifleregs.txt" For Input As #1
    Ctr = 0
        Do While Not EOF(1)
            Ctr = Ctr + 1
            Input #1, rifleregs(Ctr)
        Loop
Close #1



Open App.Path & "\Data\rifletips.txt" For Input As #2
    Ctr = 0
        Do While Not EOF(2)
            Ctr = Ctr + 1
            Input #2, rifletips(Ctr)
        Loop
Close #2



Open App.Path & "\Data\rifles.txt" For Input As #3
    Ctr = 0
        Do While Not EOF(3)
            Ctr = Ctr + 1
            Input #3, Guns(Ctr), Caliber(Ctr), Grain(Ctr), Energy(Ctr)
        Loop
Close #3



frmHunting.Hide
frmRifle.Show

End Sub

Private Sub cmdTIP_Click()
'this button displays two messeges one after the other that, that give the user the 2 most important tips for hunting.

MsgBox "BE SAFE AT ALL TIMES.", , "#1 RULE"
MsgBox "RULE #2: ALWAYS READ THE REGULATIONS HANDBOOK.", , "RULE #2"
cmdTIP.Enabled = False
cmdTIP.Visible = False

End Sub

Private Sub cmdTrip_Click()
'this button takes the user to a form where they can set up a trip to go hunting

frmHunting.Hide
frmTrip.Show

End Sub

