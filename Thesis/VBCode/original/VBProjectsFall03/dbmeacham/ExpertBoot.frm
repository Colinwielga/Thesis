VERSION 5.00
Begin VB.Form ExpertBoot 
   BackColor       =   &H000000C0&
   Caption         =   "Shop for boots"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton wsize 
      Caption         =   "Click here to find your boot size in women's sizes"
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Size 
      Caption         =   "Click here to find your boot size in men's sizes"
      Height          =   975
      Left            =   6120
      TabIndex        =   7
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdTech 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Text            =   "Click on a boot to get a description and price!"
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox pbxFischer 
      Height          =   3975
      Left            =   4800
      Picture         =   "ExpertBoot.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.PictureBox pbxTech 
      Height          =   2655
      Left            =   360
      Picture         =   "ExpertBoot.frx":3FAC
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox pbxBandit 
      Height          =   3615
      Left            =   240
      Picture         =   "ExpertBoot.frx":5A11
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "ExpertBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: ExpertBoot (ExpertBoot.frm)
'Author David Meacham
'Date Written: Wednesday, October 23
'Purpose of form: Allows user to find information about the boot's displayed.
                    'Allows them to search for their boot size.  As well it
                    'allows them to purchase a pair of boots, adding the cost
                    'to the running total and move on to the next form.

Option Explicit
Private Sub cmdFischer_Click()
'Adds the price of the boot to a running total and moves on to next form
sum = sum + 625
ExpertBoot.Hide
Checkout.Show
End Sub

Private Sub cmdRossi_Click()
'Adds the price of the boot to a running total and moves on to next form
sum = sum + 499
ExpertBoot.Hide
Checkout.Show
End Sub


Private Sub cmdTech_Click()
'Adds the price of the boot to a running total and moves on to next form
sum = sum + 600
ExpertBoot.Hide
Checkout.Show
End Sub


Private Sub cmdWsize_Click()

End Sub

Private Sub pbxBandit_Click()
'Displays the description and the cost of the boot.
MsgBox "This is Rossignol's Bandit B2.  They are excellent for freestyle skiing.  They are $499."
End Sub

Private Sub pbxFischer_Click()
'Displays the description and the cost of the boot.
MsgBox "This is Fischer's Somma F 9000.  An excellent all mountain boot.  They are $625."
End Sub

Private Sub pbxTech_Click()
'Displays the description and the cost of the boot.
MsgBox "This is Technica's Icon ALU Comp.  They are an excellent racing boot.  They are $600."
End Sub

Private Sub Size_Click()
'Converts the users shoe size to MONDO boot sizing by searching an array
Dim found As Boolean
Dim shoe As Single
Dim shsize(1 To 23) As Single
Dim bsize(1 To 23) As Single
found = False
Dim boot As String
boot = strPath & "Bootsizemen.txt"          'creates a path to open Bootsizemen.txt
Dim i As Integer
    Open boot For Input As #1
        For i = 1 To 23
            Input #1, shsize(i)             ''places the info in BootSizemen into an array
            Input #1, bsize(i)
        Next i
shoe = InputBox("Input your shoe size.  Example: 10.5")         'asks the user to input their shoe size
i = 0
    Do Until found Or i >= 23
        i = i + 1
        If shoe = shsize(i) Then                            'searches for their shoe size
            found = True
        End If
    Loop
        If found = True Then
            MsgBox "Your boot size is " & bsize(i)          'prints out correspinding boot size
        Else
            MsgBox "Your shoe size is not found."           'prints out message that their shoe size was not found
        End If
Close #1                                                    'closes the array
End Sub

Private Sub wsize_Click()
'Converts the users shoe size to MONDO boot sizing by searching an array
Dim found As Boolean
Dim shoe As Single
Dim shsize(1 To 23) As Single
Dim bsize(1 To 23) As Single
found = False
Dim boot As String
boot = strPath & "Bootsizewomen.txt"            'creates a path to open Bootsizewomen.txt
Dim i As Integer
    Open boot For Input As #1
        For i = 1 To 17
            Input #1, shsize(i)                 'places the info in BootSizewomen into an array
            Input #1, bsize(i)
        Next i
shoe = InputBox("Input your shoe size.  Example: 10.5")     'asks the user to input their shoe size
i = 0
    Do Until found Or i >= 17
        i = i + 1
        If shoe = shsize(i) Then                            'searches for their shoe size
            found = True
        End If
    Loop
        If found = True Then
            MsgBox "Your boot size is " & bsize(i)          'prints out correspinding boot size
        Else
            MsgBox "Your shoe size is not found."           'prints out message that their shoe size was not found
        End If
Close #1                                                    'closes the array
End Sub
