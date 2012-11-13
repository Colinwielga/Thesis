VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0000FFFF&
   Caption         =   "Menu Planning"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H000000FF&
      Caption         =   "Sort the items by calorie amount."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H000080FF&
      Caption         =   "Reset my menu"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main Screen"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   2655
   End
   Begin VB.PictureBox picResults3 
      BackColor       =   &H000080FF&
      Height          =   4095
      Left            =   7320
      ScaleHeight     =   4035
      ScaleWidth      =   4155
      TabIndex        =   9
      Top             =   2640
      Width           =   4215
   End
   Begin VB.PictureBox picresults2 
      BackColor       =   &H0000C000&
      Height          =   975
      Left            =   7680
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H0000C000&
      Caption         =   "Calculate total calories"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddtomenu 
      BackColor       =   &H000080FF&
      Caption         =   "Add to my menu"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FFFF&
      Height          =   6015
      Left            =   360
      ScaleHeight     =   5955
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Load Menu Options"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblMenu 
      BackColor       =   &H0000FFFF&
      Caption         =   "Menu Choices"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblEnterItem 
      BackColor       =   &H0000FFFF&
      Caption         =   "<== Enter the corresponding number of your food item of choice"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCalories 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "# of Calories"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Food Item"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Bon Appetit:Menu Planner
'Form name: Menu (frmMenu.frm)
'Authors: Sarah Biro and Heather Denike
'Date written: 3/13/2008
'Objective: This form shows users a list of various food items to choose from.
            'Users can choose which items to add to a personal menu. The program
            'will then calculate the total amount of calories in the items chosen
            'and determine whether or not the amount is in their recommended range.
Option Explicit
Dim Item(1 To 25) As String, calories(1 To 25) As Single, CTR As Integer
Dim total As Integer
Dim number(1 To 25) As Integer



Private Sub cmdAddtomenu_Click()
Dim food As Integer, pos As Integer, found As Boolean

found = False
pos = 0

    food = txtItem.Text 'user enters the number that corresponds with food choice
    
    Do While ((Not found) And (pos < CTR)) 'exhaustive search to find the food item that corresponds with the input # from the user
        pos = pos + 1
        If number(pos) = food Then
            picResults3.Print Item(pos), Tab(45); calories(pos) 'prints the selected item and its caloric worth
            found = True
            total = total + calories(pos) 'adds calories from selected item to the total caloric intake for later use
        End If
    Loop




End Sub

Private Sub cmdcalculate_Click()


   
    
    
    If total > High Then 'taken from previous calculation
        MsgBox ("Sorry, you have exceeded your recommended daily caloric intake.")
    ElseIf total < low Then MsgBox ("Sorry, you have not reached the minimum recommended daily caloric intake.")
    Else:
            picresults2.Print "Eating these items, "
            picresults2.Print "you would consume"
            picresults2.Print FormatNumber(total, 1); " calories."
            'prints results if user selected foods that fall within their expected range
    End If
End Sub

Private Sub cmdLoad_Click()

    CTR = 0
    
    
    Open App.Path & "\menu.txt" For Input As #1
        'opens file
    
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, number(CTR), Item(CTR), calories(CTR)
        picResults.Print number(CTR); ") "; Item(CTR), Tab(50); calories(CTR)
            'prints number, menu item, and their caloric worth
    Loop




End Sub

Private Sub cmdMain_Click()
    frmmain.Show
    frmMenu.Hide
    'returns to main form

End Sub

Private Sub cmdReset_Click()
'resets the menu in case a person has gone over their expected calorie range
    picResults3.Cls
    total = 0
End Sub

Private Sub cmdSort_Click()
Dim pass As Integer, tempCalories As String, tempItem As String
Dim j As Integer, k As Integer, tempNumber As Integer
'clears screen from load
    picResults.Cls

'bubble sort food options in descending caloric order
    For pass = 1 To CTR - 1
        For j = 1 To CTR - pass
            If calories(j) < calories(j + 1) Then
                tempCalories = calories(j)
                calories(j) = calories(j + 1)
                calories(j + 1) = tempCalories
                tempItem = Item(j)
                Item(j) = Item(j + 1)
                Item(j + 1) = tempItem
                tempNumber = number(j)
                number(j) = number(j + 1)
                number(j + 1) = tempNumber
                
            End If
        Next j
    Next pass

'prints options in descending caloric order
    For k = 1 To CTR
        picResults.Print number(k); ") "; Item(k); Tab(50); calories(k)
    Next k
        
End Sub
