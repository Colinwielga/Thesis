VERSION 5.00
Begin VB.Form Huaiyang_cuisine 
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   Picture         =   "Hunan_cuisine.frx":0000
   ScaleHeight     =   9690
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   1095
      Left            =   12000
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdintroduce 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Introduce Huaiyang cuisine"
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Type the name of the dish for the picture"
      Height          =   1215
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtname 
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   7920
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800080&
      Height          =   5655
      Left            =   5400
      ScaleHeight     =   5595
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
   End
   Begin VB.CommandButton cmdsortb 
      BackColor       =   &H0000C000&
      Caption         =   "List all the dishes by Lenth"
      Height          =   1215
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdsorta 
      BackColor       =   &H0000FF00&
      Caption         =   "list all the dishes A-Z"
      Height          =   1335
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Huaiyang_cuisine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chinese food
'Form Name: huaiyang
'Author: Yifan Pang
'Date Written: feb 23 2010
'The purpose of this form is introduce huaiyang dishes
Option Explicit
Dim I As Integer

Private Sub cmdintroduce_Click()
Dim huaiyang(1 To 100) As String   'read and show information
Dim ctr As Integer
Dim n As Integer
picResults.Cls
Open App.Path & "\huaiyang.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, huaiyang(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picResults.ForeColor = RGB(0, 0, 0)
    picResults.Print huaiyang(n)
Next n
End Sub

Private Sub cmdReturn_Click()
Huaiyang_cuisine.Hide
China.Show

End Sub

Private Sub cmdsorta_Click()
Dim pass As Integer, c As Single, temp As String, x As Integer 'sort by A-Z
picResults.Cls
For pass = 1 To I - 1
        For c = 1 To I - pass
        If dishes(c) > dishes(c + 1) Then
            temp = dishes(c)
            dishes(c) = dishes(c + 1)
            dishes(c + 1) = temp

        End If
    Next c
Next pass
For x = 1 To I
picResults.Print dishes(x)
 Next x

End Sub

Private Sub cmdsortb_Click()
Dim pass As Integer, c As Single, temp As String, x As Integer 'use len function for sort
picResults.Cls
For pass = 1 To I - 1
        For c = 1 To I - pass
        If Len(dishes(c)) > Len(dishes(c + 1)) Then
            temp = dishes(c)
            dishes(c) = dishes(c + 1)
            dishes(c + 1) = temp

        End If
    Next c
Next pass
For x = 1 To I
picResults.Print dishes(x)
 Next x
   
End Sub

Private Sub Command1_Click() 'use if then statement to shoe pictures
Dim name As String
picResults.Print
name = txtname.Text
If name = "Duck Egg and Pork Porridge" Then
picResults.picture = LoadPicture(App.Path & "\pidan.jpg")
ElseIf name = "Sour Vegetable Fish Pot" Then
picResults.picture = LoadPicture(App.Path & "\yu.jpg")
ElseIf name = "Pot Stickers" Then
picResults.picture = LoadPicture(App.Path & "\jiaozi.jpg")
ElseIf name = "Pork and Shrimp Dumpling Noodles" Then
picResults.picture = LoadPicture(App.Path & "\xiaozi.jpg")
ElseIf name = "Giant lion’s head Meatball" Then
picResults.picture = LoadPicture(App.Path & "\shizi.jpg")
ElseIf name = "Steamed Pork Rice Wraps" Then
picResults.picture = LoadPicture(App.Path & "\saomai.jpg")
Else
picResults.picture = LoadPicture(App.Path & "\error.jpg")
End If




End Sub

Private Sub Form_Load() ' auto load txt file

I = 0
Open App.Path & "\dishes.txt" For Input As #1

Do While Not EOF(1)
    I = I + 1
    Input #1, dishes(I)
Loop
Close #1
End Sub

