VERSION 5.00
Begin VB.Form frmDescription1 
   BackColor       =   &H00000000&
   Caption         =   "Important Information"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FF00&
      Caption         =   "Back to the case files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9600
      Width           =   2295
   End
   Begin VB.PictureBox picdisorg 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   7920
      ScaleHeight     =   8055
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   1680
      Width           =   7215
   End
   Begin VB.CommandButton cmdDIS 
      BackColor       =   &H000080FF&
      Caption         =   "Click to read a description of Organized and Disorganized offenders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9840
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picRape 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   240
      ScaleHeight     =   8055
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   1680
      Width           =   7455
   End
   Begin VB.CommandButton cmdrape 
      BackColor       =   &H000080FF&
      Caption         =   "Click to read the four rapist typlogies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmDescription1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
'This takes the user back to case files selection form
    frmDescription1.Hide
    frmCasefiles.Show
    
End Sub

Private Sub cmdDIS_Click()
'I delcare these two variables for this button
    Dim pos As Integer, ctr As Integer
    
    picdisorg.Cls 'Clear my picture boxes of old data
    
    'open a txt file i made and calling the data #3
        Open App.Path & "\orgdis.txt" For Input As #3
        Do Until EOF(3) 'I want to circle until i have reached the end of the data in the file
            pos = pos + 1 'This keeps track of how many times i actually did the loop
            Input #3, orgdis(pos) 'i have the data stored in here now
        Loop 'I loop until i read the file and have it now all saved as orgdis
        Close #3 'I close this file as to not have it interfer with other files
        
    For ctr = 1 To 22 'The range of my file i want it to read
        picdisorg.Print orgdis(ctr) 'print the results
    Next ctr 'until the end of the ctr amount
    
End Sub

Private Sub cmdrape_Click()
'I declare these two variables for this button
Dim pos As Integer, ctr As Integer

picRape.Cls 'Clear my picture boxes of old data
 'open a txt file i made and calling the data #2
    Open App.Path & "\rapisttypes.txt" For Input As #2
    Do Until EOF(2) 'I want to circle until i have reached the end of the data in the file
        pos = pos + 1 'This keeps track of how many times i actually did the loop
        Input #2, Rapetypes(pos) 'i have the data stored in here now
    Loop 'I loop until i read the file and have it now all saved as Rapetypes
    Close #2 'I close this file as to not have it interfer with other files
    
For ctr = 1 To 26 'The range of my file i want it to read
    picRape.Print Rapetypes(ctr) 'print the results
Next ctr 'until the end of the ctr amount
    
End Sub

