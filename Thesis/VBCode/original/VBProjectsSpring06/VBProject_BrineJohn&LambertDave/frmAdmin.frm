VERSION 5.00
Begin VB.Form frmAdmin 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Go back to the Navigate page"
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Click to print totals for the day"
      Height          =   1095
      Left            =   2280
      TabIndex        =   1
      Top             =   5280
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   480
      ScaleHeight     =   4275
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dave and JB's ski and skate store
'frmFront
'Dave Lambert and John Brine
'Thursday March 23
'This form is used by the administator only, it is used to show the total of skis and skates purchased

Private Sub cmdNavigate_Click()
    'moves back to the navigate page
    frmNavigate.Visible = True
    frmAdmin.Visible = False
End Sub

Private Sub cmdPrint_Click()
    'establishes all variables needed for subroutine
    Dim pos As Integer
    Dim skatetotals(1 To 100) As Integer
    Dim skitotals(1 To 100) As Integer
    
    picResults.Cls 'clears the picture box
    
    
    'prints headers
    picResults.Print "ski total for the day"
    picResults.Print "********************************************************************************************************************************************"
    
    'opens the file from text file
    Open App.Path & "\skitotals2.txt" For Input As #3
        Do Until EOF(3)
        pos = pos + 1
        Input #3, skitotals(pos)
        'prints the file
        picResults.Print skitotals(pos)
    Loop
    Close #3 'closes the file
    
    'prints the header
    picResults.Print "skate totals for the day"
    picResults.Print "*****************************************************************************************************************************************"
    
    'opens the file
    Open App.Path & "\skatetotals2.txt" For Input As #4
        pos = 0
        Do Until EOF(4)
        pos = pos + 1
        Input #4, skatetotals(pos)
        
        'prints the file
        picResults.Print skatetotals(pos)
    Loop
    Close #4
End Sub

