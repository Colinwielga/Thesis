VERSION 5.00
Begin VB.Form frmPhase2 
   BackColor       =   &H80000006&
   Caption         =   "Pole Vaulting pg 2"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   Picture         =   "Phase2.frx":0000
   ScaleHeight     =   8220
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H80000007&
      Height          =   420
      Left            =   8100
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7830
      Width           =   1770
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H000000FF&
      Caption         =   "History of Pole vaulting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   0
      Picture         =   "Phase2.frx":DD45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5805
      Width           =   1770
   End
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H000000FF&
      Caption         =   "Phase 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6750
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   6105
      Left            =   5280
      ScaleHeight     =   302.25
      ScaleMode       =   2  'Point
      ScaleWidth      =   297.75
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdAvg 
      BackColor       =   &H000000FF&
      Caption         =   "Average height of All MIAC   Pole Vaulters"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5265
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6345
      Width           =   2295
   End
   Begin VB.CommandButton cmdMIAC 
      BackColor       =   &H000000FF&
      Caption         =   "Top Pole vaulters of the MIAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2430
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6345
      Width           =   2295
   End
End
Attribute VB_Name = "FrmPhase2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pole Vaulting'
'Phase 2
'Aaron Laine
'10/19/09
'This Phase allows one to go to a website and look up the History of the Pole Vault. It also Shows the Pole Vaulters of the MIAC and how they Rank, and then computes the average of those Vaulter. After that it sends you to Phase 3.'
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Dim runningtotal As Integer

Private Sub cmdAvg_Click()
' variables used'
  Dim J As Integer, Avg As Single
    
    'calculate and print the average height by all Vaulters'
    Avg = runningtotal / ctr
    picResults.Print "Average Height of Vaulters = "; "   "; FormatNumber(Avg, 3)
    picResults.Print "**************************************"
    
    'print the Vaulters who have gained more than the average height'
    picResults.Print "The above average Pole Vaulters are:"
    picResults.Print "Vaulter", "Height"
    
    For J = 1 To ctr
        If metersArray(J) > Avg Then
             picResults.Print Vaulter(J); "    "; Verticle(J)
        End If
    Next J
    
    Close #1        'Close the file used for input
    
    'disable the command button for printing above average Vaulters

End Sub
'Opens up a website'
Private Sub cmdHistory_Click()
ShellExecute Me.hwnd, "open", "http://trackfieldevents.com/history/pole-vault-history", "", "", SW_SHOW
End Sub
'An array to sort the Vaulters into first to last place'
Private Sub cmdMIAC_Click()
Dim meters As Single
Dim vault, vert As String
ctr = 0

Open App.Path & "\PoleVault.txt" For Input As #1

'Print Vaulter and height'
picResults.Print "Vaulter"; "   "; "Height (FT)"; "   "; "Meters"
picResults.Print "------------------------------------------------"
    Do Until EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        ctr = ctr + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, vault, vert, meters
        Vaulter(ctr) = vault
        Verticle(ctr) = vert
        metersArray(ctr) = meters
        picResults.Print Vaulter(ctr); "    "; Verticle(ctr); "   "; metersArray(ctr)
        
        'add feet to runningTotal
        runningtotal = runningtotal + meters
    Loop
    Close #1        'Close the file used for input'
   
    
     'disable the command button for Reading the file'
    cmdMIAC.Enabled = False
    
    'enable the command button for printing the above average runners'
    cmdAvg.Enabled = True
    
    'Enable Switch button
    cmdSwitch.Enabled = True

End Sub

'Hides phase2 and shows Phase 1

Private Sub cmdMove_Click()
frmPhase1.Show
FrmPhase2.Hide
End Sub
'hides Phase 2 and shows phase 3
Private Sub cmdSwitch_Click()
frmPhase3.Show
FrmPhase2.Hide
End Sub


