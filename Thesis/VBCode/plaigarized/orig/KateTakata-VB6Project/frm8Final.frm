VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H000040C0&
   Caption         =   "Congratulations!"
   ClientHeight    =   11370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   11370
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead2 
      Caption         =   "Read and Display Second Electric's Information"
      Height          =   735
      Left            =   5880
      TabIndex        =   6
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton cmdRead1 
      Caption         =   "Read and Display First Electric's Information"
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   10200
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   1320
      ScaleHeight     =   7515
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   840
      Width           =   8055
   End
   Begin VB.CommandButton cmdReadC 
      Caption         =   "Read and Display Catwalk's Infomation"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   10200
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Label lblSchedule 
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm8Final.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   2
      Top             =   8640
      Width           =   9255
   End
   Begin VB.Label lblTitle1 
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations! You've successfully created a light plot!"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the final form in the program.
    Dim instChoose(1 To 100) As String
    Dim gelChoose(1 To 100) As String
    Dim CTR As Integer

Private Sub cmdReadC_Click()
'This reads the file made about the catwalk and prints it into the picture box.
    Dim pos As Integer
    
    picResults.Print "Electric #", "Position", "Instrument", "Gel Color"
    picResults.Print "*********************************************************************************************"
      
    Open App.Path & "\Catwalk.txt" For Input As #1
        
        Do Until EOF(1)
            CTR = CTR + 1
            Input #1, CTR, instChoose(CTR), gelChoose(CTR)
        Loop
    
    For pos = 1 To CTR
        picResults.Print "Catwalk", pos, instChoose(pos), gelChoose(pos)
    Next pos
    
    Close #1
    
End Sub

Private Sub cmdRead1_Click()
'This reads the file made about the first electric and prints it into the picture box.
    Dim pos As Integer
    
    Open App.Path & "\First.txt" For Input As #1
        
        Do Until EOF(1)
            CTR = CTR + 1
            Input #1, CTR, instChoose(CTR), gelChoose(CTR)
        Loop
        
        For pos = 1 To CTR
            picResults.Print "First Electric", pos, instChoose(pos), gelChoose(pos)
        Next pos
        
        Close #1
        
End Sub

Private Sub cmdRead2_Click()
'This reads the file made about the second electric and prints it into the picture box.
    Dim pos As Integer
    
    Open App.Path & "\Second.txt" For Input As #1
        
        Do Until EOF(1)
            CTR = CTR + 1
            Input #1, CTR, instChoose(CTR), gelChoose(CTR)
        Loop
        
        For pos = 1 To CTR
            picResults.Print "Second Electric", pos, instChoose(pos), gelChoose(pos)
        Next pos
        
        Close #1

End Sub

Private Sub cmdQuit_Click()
'Ends the program.
    End
End Sub
