VERSION 5.00
Begin VB.Form BeginnerBoot 
   BackColor       =   &H00800080&
   Caption         =   "Shop for boots"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton msize 
      Caption         =   "Click here to find your boot size in men's sizes"
      Height          =   855
      Left            =   7800
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton wsize 
      Caption         =   "Click here to find your boot size in women's sizes"
      Height          =   855
      Left            =   5280
      TabIndex        =   7
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdTechnica 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdFischer 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdRossi 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000C000&
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Text            =   "Click on a boot for a description and price!"
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox pbxFishcer 
      Height          =   2895
      Left            =   6600
      Picture         =   "BeginnerBoot.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
   End
   Begin VB.PictureBox pbxTechnica 
      Height          =   2655
      Left            =   2520
      Picture         =   "BeginnerBoot.frx":2116
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
   Begin VB.PictureBox pbxRossi 
      Height          =   3495
      Left            =   600
      Picture         =   "BeginnerBoot.frx":3CDB
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "BeginnerBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DavesSkiShop (Ski.vbp)
'Form name: BeginnerBoot (BeginnerBoot.frm)
'Author David Meacham
'Date Written: Wednesday, October 23
'Purpose of form: Allows user to find information about the boot's displayed.
                    'Allows them to search for their boot size.  As well it
                    'allows them to purchase a pair of boots, adding the cost
                    'to the running total and move on to the next form.

Option Explicit

Private Sub cmdFishcer_Click()
'Displays a description and price of the boot and moves on to next form
MsgBox "This is Fischer's Somma F 7000.  A very responsive boot.  They are $450."
End Sub

Private Sub cmdFischer_Click()
'Adds the price of the boot to the running total
sum = sum + 450
BeginnerBoot.Hide
Checkout.Show
End Sub

Private Sub cmdRossi_Click()
'Adds the price of the boot to the running total and moves on to next form
sum = sum + 300
BeginnerBoot.Hide
Checkout.Show
End Sub

Private Sub cmdTechnica_Click()
'Adds the price of the boot to the running total and moves on to next form
sum = sum + 375
BeginnerBoot.Hide
Checkout.Show
End Sub

Private Sub msize_Click()
'Converts the users shoe size to MONDO boot sizing by searching an array
Dim found As Boolean
Dim shoe As Single
Dim shsize(1 To 23) As Single
Dim bsize(1 To 23) As Single
Dim boot As String
boot = strPath & "Bootsizemen.txt"                        'places the info in BootSizemen into an array
found = False
Dim i As Integer
    Open boot For Input As #1
        For i = 1 To 23
            Input #1, shsize(i)
            Input #1, bsize(i)
        Next i
shoe = InputBox("Input your shoe size.  Example: 10.5")     'asks the user to input their shoe size
i = 0
    Do Until found Or i >= 23
        i = i + 1
        If shoe = shsize(i) Then                        'searches for their shoe size
            found = True
        End If
    Loop
        If found = True Then
            MsgBox "Your boot size is " & bsize(i)      'prints out correspinding boot size
        Else
            MsgBox "Your shoe size is not found."       'prints out message that their shoe size was not found
        End If
Close #1                                                'closes the array
End Sub

Private Sub pbxFishcer_Click()
'Displays a description and price of the boot
MsgBox "This is Fischer's Somma 7000.  A very responsive boot.  They are $450."
End Sub

Private Sub pbxRossi_Click()
'Displays a description and price of the boot
MsgBox "This is Rossignol's Axium SX.  An extremely comfertable boot.  They are $300."
End Sub

Private Sub pbxTechnica_Click()
'Displays a description and price of the boot
MsgBox "This is Technica's Gamma 9.  A very comfertable boot.  They are $375."
End Sub


Private Sub wsize_Click()
'Converts the users shoe size to MONDO boot sizing by searching an array
Dim found As Boolean
Dim shoe As Single
Dim shsize(1 To 23) As Single
Dim bsize(1 To 23) As Single
Dim boot As String
boot = strPath & "Bootsizewomen.txt"        'places the info in BootSizewomen into an array
found = False
Dim i As Integer
    Open boot For Input As #1
        For i = 1 To 17
            Input #1, shsize(i)             'places the info in BootSizeWomen into an array
            Input #1, bsize(i)
        Next i
shoe = InputBox("Input your shoe size.  Example: 10.5")    'asks the user to input their shoe size
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
