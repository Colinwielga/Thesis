VERSION 5.00
Begin VB.Form IdzerdaHouseChores 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton display_names 
      Caption         =   "Display Names"
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton quit_cmd 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton clear_cmd 
      Caption         =   "Clear Results"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox results 
      BackColor       =   &H00FFFF00&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   4755
      TabIndex        =   7
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton findhousemate_cmd 
      Caption         =   "Find Housemate"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox numberbox 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton findchore_cmd 
      BackColor       =   &H00FFFF00&
      Caption         =   "Find Chore"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox namebox 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C000C0&
      Caption         =   "11) Pick up and vacuum mustard room"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   4095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C000C0&
      Caption         =   "10) Pick up and vaccuum dining room, wipe off tables"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C000C0&
      Caption         =   "9) Vacuum downstairs"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   4095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C000C0&
      Caption         =   "8) Sweep and mop stairs"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C000C0&
      Caption         =   "7) Pick up and vacuum living room"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   4095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C000C0&
      Caption         =   "6) Week off"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C000C0&
      Caption         =   "5) Clean kitchen counters and microwave"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C000C0&
      Caption         =   "4) Sweep and mop entryway"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C000C0&
      Caption         =   "3) Take out trash and recycling"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C000C0&
      Caption         =   "2) Clean computer room, laundry room, and solarium"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C000C0&
      Caption         =   "1) Sweep and mop kitchen"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C000C0&
      Caption         =   "Numeric Chore Codes"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C000C0&
      Caption         =   "Enter a Chore Number"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Caption         =   "-or-"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Enter a Housemate's Name"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "IdzerdaHouseChores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JohannaHatch (IdzerdaHouseChores.vbp)
'IdzerdaHouseChores (IdzerdaHouseChores.frm)
'Author: Johanna Hatch
'11/3/03
'Purpose: to create a simple way for the eleven residents of the Idzerda House to keep track of their chores

Dim Path As String
Dim chore(1 To 11) As String
Dim housematename(1 To 11) As String


Public Sub clear_cmd_Click()
results.Cls
End Sub

Private Sub display_names_Click()
MsgBox "Johanna, Sarah, Robyn, Roxi, Kate, Keely, Jasna, Samantha, Danielle, Maggie, Kirsten", , "Housemate Names"
End Sub

Public Sub findchore_cmd_Click()

Dim n As String
Dim j As Integer
n = namebox.Text 'reads in name from textbox
j = 0
Open Path & "chores.txt" For Input As #1
For j = 1 To 11
    Input #1, chore(j), housematename(j)
Next j
Do Until j <= 11 'search through the array to find the name, stopping when it reaches 11
    j = j + 1
    If n = housematename(j) Then 'when the text matches a name in the array, it will move onto the next step
        results.Print housematename(j); "'s chore is "; chore(j) 'prints name and corresponding chore
    End If
Loop
 

Close #1

End Sub

Private Sub findhousemate_cmd_Click()
Dim Found As Boolean
Dim num As Integer
Dim j As Integer
num = numberbox.Text 'enter chore's number from code
j = 0
NotFound = True
Open Path & "chores.txt" For Input As #1
For j = 1 To 11
    Input #1, chore(j), housematename(j)
Next j

results.Print housematename(num); "'s chore is "; chore(num) 'prints the chore that matches the entered number and corresponding name

Close #1
End Sub

Public Sub Form_Load()

Path = "m:/CS130/JohannaHatch/"
End Sub

Private Sub Option1_Click()

End Sub

Public Sub quit_cmd_Click()
End
End Sub
