VERSION 5.00
Begin VB.Form frmGear 
   Caption         =   "Bryan Mills"
   ClientHeight    =   7950
   ClientLeft      =   2790
   ClientTop       =   1305
   ClientWidth     =   9720
   ForeColor       =   &H80000018&
   LinkTopic       =   "Form1"
   Picture         =   "frmGear.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   9720
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Box"
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox picOutput 
      Height          =   4695
      Left            =   6720
      ScaleHeight     =   4635
      ScaleWidth      =   2715
      TabIndex        =   28
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdMap 
      Caption         =   "Get The Lake Map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      TabIndex        =   13
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdBaitBucket 
      Caption         =   "Bait Bucket"
      Height          =   615
      Left            =   7080
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdTackelBox 
      Caption         =   "TackleBox"
      Height          =   615
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Jitter Bug"
      Height          =   735
      Left            =   3600
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdCamera 
      Caption         =   "Camera"
      Height          =   735
      Left            =   3600
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSpoon 
      Caption         =   "Daredevil"
      Height          =   735
      Left            =   3720
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdSpinner 
      Caption         =   "Spinner Bait"
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdNet 
      Caption         =   "Landing Net"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdReel 
      Caption         =   "Reel"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdLeeche 
      Caption         =   "Leeches"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdMinnow 
      Caption         =   "Minnows"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdWorm 
      Caption         =   "Worm"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdRapala 
      Caption         =   "Rapala"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRod 
      Caption         =   "Fishing Rod"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblBudget 
      BackColor       =   &H0000FFFF&
      Caption         =   "Budget"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   29
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblBox 
      Caption         =   "20.00"
      Height          =   255
      Left            =   7800
      TabIndex        =   27
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblBucket 
      Caption         =   "5.00"
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblSpinner 
      Caption         =   "3.00"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label lblDaredevil 
      Caption         =   "2.30"
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblCamera 
      Caption         =   "3.00"
      Height          =   255
      Left            =   5520
      TabIndex        =   23
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblBug 
      Caption         =   "3.00"
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   8520
      TabIndex        =   21
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblNet 
      Caption         =   "37.50"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label lblReel 
      Caption         =   "40.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lblLeech 
      Caption         =   "4.75"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblMinnow 
      Caption         =   "3.50"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblWorm 
      Caption         =   "2.50"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblRapala 
      Caption         =   "3.50"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblRod 
      Caption         =   "25.99"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmGear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project is a bass fishing guideing service (final.frm)
'This is the buying for for gear (gear.frm)
'Bryan Mills
'March 24, 2006
'This form will calculte the remaining money in your budget after purchasing items


Option Explicit
Dim rod, reel, spinner, spoon, baitbucket, bug, camera, leech, minnow, net, rapala, tacklebox, worm, total As Currency
Dim output As Currency
Dim newbudget As Single
Dim I As Integer
Dim found As Boolean
Dim counter As Single
'this defines all of the variables



Private Sub cmdLoad_Click()
    picOutput.Print newbudget
    'this prints the Budget you can spend from in the beginning of the pic box
End Sub


Private Sub cmdBaitBucket_Click()
        found = False
        I = 1
        Do While Not found
            If gear(I) = "Bait Bucket" Then
                newbudget = newbudget - price(I)
            found = True
        End If
        I = I + 1
        If newbudget < 0 Then
            MsgBox "You are in Debt. Stop Buying!"
        End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdBug_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Jitter Bug" Then
           newbudget = newbudget - price(I)
        found = True
    End If
    I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
        picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdCamera_Click()
        found = False
        I = 1
        Do While Not found
            If gear(I) = "Camera" Then
               newbudget = newbudget - price(I)
            found = True
        End If
        I = I + 1
        If newbudget < 0 Then
            MsgBox "You are in Debt. Stop Buying!"
        End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdClear_Click()
    picOutput.Cls
    'this clears the pic box
End Sub

Private Sub cmdLeeche_Click()
        found = False
        I = 1
        Do While Not found
            If gear(I) = "Leeches" Then
                newbudget = newbudget - price(I)
            found = True
        End If
         I = I + 1
        If newbudget < 0 Then
            MsgBox "You are in Debt. Stop Buying!"
        End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub
Private Sub cmdMap_Click()
    Dim pos As Single
        pos = 1
    Open App.Path & "\HotSpots.txt" For Input As #2
        Do Until EOF(2)
            Input #2, location(pos), method(pos), bait(pos)
            pos = pos + 1
        Loop
    Close #2
    'this reads in the second array which contains the hot spots in string form
    frmGear.Hide
    frmLake.Show
    'this takes you to the next form
End Sub

Private Sub cmdMinnow_Click()
        found = False
        I = 1
        Do While Not found
            If gear(I) = "Minnows" Then
                newbudget = newbudget - price(I)
            found = True
        End If
        I = I + 1
        If newbudget < 0 Then
            MsgBox "You are in Debt. Stop Buying!"
        End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdNet_Click()
        found = False
        I = 1
        Do While Not found
            If gear(I) = "Landing Net" Then
               newbudget = newbudget - price(I)
            found = True
        End If
         I = I + 1
        If newbudget < 0 Then
            MsgBox "You are in Debt. Stop Buying!"
        End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub
Private Sub cmdRapala_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Rapala" Then
            newbudget = newbudget - price(I)
        found = True
    End If
     I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub
Private Sub cmdReel_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Reel" Then
            newbudget = newbudget - price(I)
        found = True
    End If
    I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub


Private Sub cmdRod_Click()
            found = False
            I = 1
            Do While Not found
                If gear(I) = "Fishing Rod" Then
                    newbudget = newbudget - price(I)
                found = True
            End If
            'this do while loop is looking for a string in the array and once found outputs the remaining budget
            I = I + 1
            If newbudget < 0 Then
                MsgBox "You are in Debt. Stop Buying!"
            End If
        'this if statement looks at the budget in comparison with zero and if less than zero then a message box pops up
        Loop
        
           picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
           'if the remaining budget is greater than zero then the number is printed in a pic box
End Sub

Public Sub loadbudget()
    newbudget = budget
    picOutput.Print "Money to spend="; FormatCurrency(newbudget)
    'this loads the new budget from the old budget and formats it so that it is in the appropriate form for money
End Sub

Private Sub cmdSpinner_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Spinner Bait" Then
            newbudget = newbudget - price(I)
        newbudget = newbudget - price(I)
        found = True
    End If
     I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdSpoon_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Daredevil" Then
          newbudget = newbudget - price(I)
        found = True
    End If
     I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdTackelBox_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Tackle Box" Then
           newbudget = newbudget - price(I)
        found = True
    End If
     I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

Private Sub cmdWorm_Click()
    found = False
    I = 1
    Do While Not found
        If gear(I) = "Worm" Then
           newbudget = newbudget - price(I)
        found = True
    End If
    I = I + 1
    If newbudget < 0 Then
        MsgBox "You are in Debt. Stop Buying!"
    End If
    Loop
       picOutput.Print "Total Remaining ="; FormatCurrency(newbudget)
End Sub

