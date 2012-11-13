VERSION 5.00
Begin VB.Form frmweightclass 
   Caption         =   "Weight Classes"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "Weight Classes Form.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults2 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   3465
      TabIndex        =   13
      Top             =   6840
      Width           =   3525
   End
   Begin VB.PictureBox picresults 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   3435
      TabIndex        =   12
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtweightclass 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   6
      Top             =   7080
      Width           =   3735
   End
   Begin VB.CommandButton cmdshowtitleholder 
      Caption         =   "Show Title Holder for Selected Weight Class"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   5
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.OptionButton optkg 
      Caption         =   "Kilograms (kg)"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton optlb 
      Caption         =   "Pounds (lb)"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblinstruction 
      Caption         =   "Please choose  one of the weight classes from above and write it in the text box.  Then press the button."
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   14
      Top             =   8160
      Width           =   3615
   End
   Begin VB.Label lblheavyweight 
      Caption         =   "Heavyweight"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblmiddleweight 
      Caption         =   "Middleweight"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblwelterweight 
      Caption         =   "Welterweight"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lbllightheavyweight 
      Caption         =   "Light Heavyweight"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label lbllightweight 
      Caption         =   "Lightweight"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lbldirection 
      Caption         =   $"Weight Classes Form.frx":22B4D
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmweightclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclear_Click()
picresults2.Cls
picresults.Picture = Nothing

End Sub

Private Sub cmdgoback_Click()
frmweightclass.Hide
frmmainscreen.Show
End Sub



Private Sub cmdshowtitleholder_Click()
Dim weightclass(1 To 100) As String, classinput As String, ctr As Integer
Dim fighter(1 To 100) As String, ctr2 As Integer
    ctr = 0
    classinput = txtweightclass.Text
    'defines what is written in the text box

Open App.Path & "\WeightClasses.txt" For Input As #1 'opening file
Do While Not EOF(1)
    ctr = ctr + 1
Input #1, weightclass(ctr), fighter(ctr) 'telling the program what is in the file
If (StrComp(weightclass(ctr), classinput, vbTextCompare) = 0) Then 'compares two strings (text and file), gives 0 if both strings match. saving position of matched fighter name
      ctr2 = ctr 'telling program what to print based on what was put into the text box
        If classinput = "Lightweight" Then
            picresults.Picture = LoadPicture(App.Path + "\frankieedgar.jpg")
            picresults2.Print fighter(ctr2)
        ElseIf classinput = "Welterweight" Then
            picresults.Picture = LoadPicture(App.Path + "\georgesstpierre.jpg")
            picresults2.Print fighter(ctr2)
        ElseIf classinput = "Middleweight" Then
            picresults.Picture = LoadPicture(App.Path + "\andersonsilva.jpg")
            picresults2.Print fighter(ctr2)
        ElseIf classinput = "Light Heavyweight" Then
             picresults.Picture = LoadPicture(App.Path + "\mauriciorua.jpg")
             picresults2.Print fighter(ctr2)
        ElseIf classinput = "Heavyweight" Then
            picresults.Picture = LoadPicture(App.Path + "\brocklesnar.jpg")
            picresults2.Print fighter(ctr2)
        Else
        
         MsgBox "Invalid Entry"
        End If
Else

End If

Loop
Close #1
    
    

    
End Sub

Private Sub optkg_Click()
Dim yourweight As Single, yourname As String
    yourweight = InputBox("Please type in your weight in kilograms to determine which UFC weight class you would fall into.", "Your Weight in kilograms")
Select Case yourweight 'telling computer what to put in the message box based on what the user inputed
    Case Is < 66
        MsgBox yourname & "Sorry, your weight of " & yourweight & "kg is to light to compete in the UFC!"
    Case 66 To 70
        MsgBox yourname & " At " & yourweight & "kg you would compete as a Lightweight."
    Case 71 To 77
        MsgBox yourname & " At " & yourweight & "kg you would compete as a Welterweight."
    Case 78 To 84
        MsgBox yourname & " At " & yourweight & "kg you would compete as a Middleweight."
    Case 85 To 93
        MsgBox yourname & " At " & yourweight & "kg you would compete as a Light Heavyweight."
    Case 94 To 120
        MsgBox yourname & " At " & yourweight & "kg you would compete as a Heavyweight."
    Case Is >= 120
        MsgBox yourname & " At " & yourweight & "kg you are too heavy for the UFC, but should compete in a different MMA"
    Case Else
        MsgBox yourname & " Sorry, your weight of " & yourweight & "kg was invalid.  Please try and enter a different weight."
    End Select
End Sub

Private Sub optlb_Click()
Dim yourweight As Single, yourname As String
    yourweight = InputBox("Please type in your weight in pounds to determine which UFC weight class you would fall into.", "Your Weight in pounds")
Select Case yourweight 'telling computer what to put in the message box based on what the user inputed
    Case Is < 146
        MsgBox yourname & "Sorry, your weight of " & yourweight & "lb is to light to compete in the UFC!"
    Case 146 To 155
        MsgBox yourname & " At " & yourweight & "lb you would compete as a Lightweight."
    Case 156 To 170
        MsgBox yourname & " At " & yourweight & "lb you would compete as a Welterweight."
    Case 171 To 185
        MsgBox yourname & " At " & yourweight & "lb you would compete as a Middleweight."
    Case 185 To 205
        MsgBox yourname & " At " & yourweight & "lb you would compete as a Light Heavyweight."
    Case 206 To 265
        MsgBox yourname & " At " & yourweight & "lb you would compete as a Heavyweight."
    Case Is > 265
        MsgBox yourname & " At " & yourweight & "lb you are too heavy for the UFC, but should compete in a different MMA"
    Case Else
        MsgBox yourname & " Sorry, your weight of " & yourweight & "lb was invalid.  Please try and enter a different weight."
    End Select
End Sub

