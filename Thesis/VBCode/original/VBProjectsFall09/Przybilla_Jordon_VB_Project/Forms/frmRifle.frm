VERSION 5.00
Begin VB.Form frmRifle 
   Caption         =   "Rifle Hunting"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "frmRifle.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortbyEnergy 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Rifles by Energy At 100 yards"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdStats 
      BackColor       =   &H000080FF&
      Caption         =   "Display Common Rifles"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortbyCaliber 
      BackColor       =   &H000080FF&
      Caption         =   "Sort Rifles By Caliber"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSuggestions 
      BackColor       =   &H000080FF&
      Caption         =   "Questions to ask yourself when selecting a firearm."
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdHunting 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Hunting Page"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.PictureBox picRifle 
      BackColor       =   &H000080FF&
      Height          =   7935
      Left            =   2160
      ScaleHeight     =   7875
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton cmdRestrictions 
      BackColor       =   &H000080FF&
      Caption         =   "Rifle Hunting Regulations for Minnesota"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRifle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MN Deer'
'Form Name: Rifle'
'Authors: Jordon Przybilla'
'Date Written: October 8, 2009
'this form will give the user a brief overview of hunting regulations and suggestions for rifle hunting

Option Explicit


Private Sub cmdHome_Click()
'takes the user back to the home page

frmRifle.Hide
frmHome.Show

End Sub

Private Sub cmdHunting_Click() 'takes the user to the hunting page

frmRifle.Hide
frmHunting.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRestrictions_Click() 'this will display the firearm regulations in minnesota
Dim a As Integer

picRifle.Visible = True
picRifle.Cls
picRifle.Print "The following regulations obtain to all large game animals in Minnesota, not just deer."
picRifle.Print

For a = 1 To Ctr
    picRifle.Print rifleregs(a)
    picRifle.Print
Next a



End Sub



Private Sub cmdSortbyCaliber_Click()
'this button will sort the rifles by caliber and diplay them
Dim pass As Integer, pos As Integer, tempGuns As String, tempCaliber As Single, tempGrain As Single, tempEnergy As Long, l As Integer

picRifle.Visible = True
picRifle.Cls
picRifle.Print "Rifle"; Tab(20); "Caliber"; Tab(30); "Grain of Common Bullet"; Tab(60); "Energy of Bullet at 100 yards"
picRifle.Print "***********************************************************************************************************************"

For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Caliber(pos) > Caliber(pos + 1) Then
            tempCaliber = Caliber(pos)
            Caliber(pos) = Caliber(pos + 1)
            Caliber(pos + 1) = tempCaliber
            tempGuns = Guns(pos)
            Guns(pos) = Guns(pos + 1)
            Guns(pos + 1) = tempGuns
            tempGrain = Grain(pos)
            Grain(pos) = Grain(pos + 1)
            Grain(pos + 1) = tempGrain
            tempEnergy = Energy(pos)
            Energy(pos) = Energy(pos + 1)
            Energy(pos + 1) = tempEnergy
        End If
    Next pos
Next pass

For l = 1 To Ctr
    picRifle.Print Guns(l); Tab(20); FormatNumber(Caliber(l), 3); Tab(40); Grain(l); Tab(70); Energy(l)
    picRifle.Print
Next l

picRifle.Print
picRifle.Print "The caliber of the rifle simply means the size of projectile the rifle fires."
picRifle.Print "Some rifles are better suited for short range shots through brush and some "
picRifle.Print "are more suited for long range shots through open fields."

End Sub



Private Sub cmdSortbyEnergy_Click()
' this button will sort the rifles by energy of the bullet at 100 yards
Dim pass As Integer, pos As Integer, tempGuns As String, tempCaliber As Single, tempGrain As Single, tempEnergy As Long, l As Integer

picRifle.Visible = True
picRifle.Cls
picRifle.Print "Rifle"; Tab(20); "Caliber"; Tab(30); "Grain of Common Bullet"; Tab(60); "Energy of Bullet at 100 yards"
picRifle.Print "***********************************************************************************************************************"

For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Energy(pos) > Energy(pos + 1) Then
            tempEnergy = Energy(pos)
            Energy(pos) = Energy(pos + 1)
            Energy(pos + 1) = tempEnergy
            tempGuns = Guns(pos)
            Guns(pos) = Guns(pos + 1)
            Guns(pos + 1) = tempGuns
            tempGrain = Grain(pos)
            Grain(pos) = Grain(pos + 1)
            Grain(pos + 1) = tempGrain
            tempCaliber = Caliber(pos)
            Caliber(pos) = Caliber(pos + 1)
            Caliber(pos + 1) = tempCaliber
        End If
    Next pos
Next pass

For l = 1 To Ctr
    picRifle.Print Guns(l); Tab(20); FormatNumber(Caliber(l), 3); Tab(40); Grain(l); Tab(70); Energy(l)
    picRifle.Print
Next l

picRifle.Print
picRifle.Print "If a rifle has less energy at 100 yards it is probably firing a flatter projectile that is better"
picRifle.Print "suited for close range shots and will have knock down power at those close ranges. These rifles "
picRifle.Print "are good for shots taken through brush or grass.  A rifle with a higher energy at 100 yards is "
picRifle.Print "more suited for long range shots because of a more aerodynamic bullet but will have less knock "
picRifle.Print "down power at longer ranges."

End Sub

Private Sub cmdStats_Click()
'this button will display common deer rifles, their caliber, the grain of bullet commonly used, and the energy of the bullet at 200 yards




picRifle.Visible = True
picRifle.Cls
picRifle.Print "Rifle"; Tab(20); "Caliber"; Tab(30); "Grain of Common Bullet"; Tab(60); "Energy of Bullet at 100 yards"
picRifle.Print "***********************************************************************************************************************"


For x = 1 To Ctr
    picRifle.Print Guns(x); Tab(20); FormatNumber(Caliber(x), 3); Tab(40); Grain(x); Tab(70); Energy(x)
    picRifle.Print
Next x

picRifle.Print
picRifle.Print
picRifle.Print

End Sub

Private Sub cmdSuggestions_Click() 'this will display helpful questions to determine the correct rifle for each user
Dim b As Integer

picRifle.Visible = True
picRifle.Cls
picRifle.Print "The following questions will give you a starting point for selecting a suitable deer rifle."
picRifle.Print

For b = 1 To Ctr
    picRifle.Print rifletips(b)
    picRifle.Print
Next b

picRifle.Print
picRifle.Print "These questions are only a starting point. To pick the correct rifle you should"
picRifle.Print "consult a veteran hunter or visit your local gun shop."

End Sub
