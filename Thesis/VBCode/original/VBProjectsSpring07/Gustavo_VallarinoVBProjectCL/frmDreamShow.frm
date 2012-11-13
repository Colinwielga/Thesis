VERSION 5.00
Begin VB.Form frmDreamShow 
   BackColor       =   &H8000000D&
   Caption         =   "Dream Team in the Field"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picmidb 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4200
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox picmidr 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5640
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   11
      Top             =   6240
      Width           =   855
   End
   Begin VB.PictureBox picmidl 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5640
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   10
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox picmidfw 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   7200
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.PictureBox picfw2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9120
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox picfwd1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9000
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.PictureBox picRD 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1440
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox picC2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox picC1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox picLD 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1680
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox picgoalie 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtDream 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   2520
      TabIndex        =   1
      Text            =   "Dream Team in the Field"
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      Height          =   855
      Left            =   7920
      TabIndex        =   0
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   10680
      X2              =   10080
      Y1              =   7440
      Y2              =   8040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   10080
      X2              =   10680
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   720
      Y1              =   7440
      Y2              =   8040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   600
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   5040
      Top             =   4200
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   5520
      X2              =   5520
      Y1              =   2040
      Y2              =   8040
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   3135
      Left            =   8880
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   3015
      Left            =   240
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   6015
      Left            =   240
      Top             =   2040
      Width           =   10455
   End
End
Attribute VB_Name = "frmDreamShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I created 11 different pictureboxes and a picture of each player will be printed
'I edited the screen so that it would look like a soccer field and things would make sense

Private Sub cmdMain_Click()
frmChampions.Show
frmDreamShow.Hide
End Sub

Private Sub Form_Load()
'I used the IF/Then/Elseif form to find the matching variable stored in the module
' Once the variable was found and it was defined it loaded and printed a picture of the players
'face in the assigned picture box, each picturebox had a positon assigned

If Goalie = "Oliver Kahn" Then
picgoalie.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\kahn.jpg")
ElseIf Goalie = "Iker Casillas" Then
picgoalie.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\Casillas.jpg")
End If

If LeftD = "Roberto Carlos" Then
picLD.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\robertocarlos.jpg")
ElseIf LeftD = "Lucio" Then
picLD.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\-lucio.jpg")
End If


If RightD = "Eric Abidal" Then
picRD.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\abidal.jpg")
ElseIf RightD = "Paolo Maldini" Then
picRD.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\maldini.jpg")
End If

If Centerd1 = "Fabio Cannavaro" Then
picC1.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\cannavaro.jpg")
ElseIf Centerd1 = "Rio Ferdinand" Then
picC1.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\ferdinand.jpg")
End If

If centerD2 = "John Terry" Then
picC2.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\terry.jpg")
ElseIf centerD2 = "Sergio Ramos" Then
picC2.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\ramos.jpg")
End If

If backmid = "Frank Lampard" Then
picmidb.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\lampard.jpg")
ElseIf backmid = "Genaro Gattuso" Then
picmidb.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\gattuso.jpg")
End If

If leftmid = "Cristiano Ronaldo" Then
picmidl.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\Cristiano.jpg")
ElseIf leftmid = "Robinho" Then
picmidl.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\robinho.jpg")
End If


If Rightmid = "Kaka" Then
picmidr.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\kaka.jpg")
ElseIf Rightmid = "Malouda" Then
picmidr.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\malouda.jpg")
End If

If offensive = "Ronaldinho" Then
picmidfw.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\dinho.jpg")
ElseIf offensive = "Zinedine Zidane" Then
picmidfw.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\zidane.jpg")
End If

If leftfwd = "Ronaldo" Then
picfwd1.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\ronaldo1.jpg")
ElseIf leftfwd = "Andriy Shevchenko" Then
picfwd1.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\shevchenko.jpg")
End If

If rightfwd = "Raul" Then
picfw2.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\Raulsmall.jpg")
ElseIf rightfwd = "Henry" Then
picfw2.Picture = LoadPicture("N:\CS130\Imad's handin\Gustavo_VallarinoVB Project CL\Players\henry.jpg")
End If








End Sub
