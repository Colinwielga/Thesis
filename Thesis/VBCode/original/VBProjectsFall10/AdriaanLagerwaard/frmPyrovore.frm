VERSION 5.00
Begin VB.Form frmPyrovore 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pyrovore Brood"
   ClientHeight    =   12255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18675
   LinkTopic       =   "Form2"
   ScaleHeight     =   12255
   ScaleWidth      =   18675
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Height          =   4095
      Left            =   6960
      ScaleHeight     =   4035
      ScaleWidth      =   4515
      TabIndex        =   9
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdPyrovoreTotal 
      Caption         =   "Total Points Spent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdPyrovoreTotalCls 
      Caption         =   "Clear Points Spent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdNumberPyrovoreCls 
      Caption         =   "Clear Number of Pyrovores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtPyrovore 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picPyrovore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton cmdPyrovore 
      Caption         =   "Back To Elites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   $"frmPyrovore.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "1-9, up to 3 per Brood: 45 pts/Each"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Pyrovore Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmPyrovore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdNumberPyrovoreCls_Click()
txtPyrovore = 0
PyrovoreTotal = PyrovoreTotal - (45 * NumberPyrovore)

End Sub

Private Sub cmdPyrovore_Click()
frmPyrovore.Hide
frmElites.Show
picPyrovore.Cls

End Sub
Public Sub loadData()
Dim j As Integer, found As Boolean


 For j = 1 To CTR
        If names(j) = "Pyrovore" Then
            picPyrovore.Print "WS     "; "BS      "; "S       "; "T      "; "W       "; "I      "; "A      "; "Ld      "; "Sv      "
            picPyrovore.Print WS(j); "       "; BS(j); "       "; S(j); "     "; T(j); "     "; W(j); "     "; I(j); "    "; A(j); "    "; Ld(j); "     "; Sv(j)
            found = True
        End If
    Next j
    
    If Not found Then
        picPyrovore.Print "Sorry."
    End If

    
End Sub

Private Sub cmdPyrovoreTotal_Click()
Dim NumberPyrovore As Single
NumberPyrovore = txtPyrovore.Text
PyrovoreTotal = PyrovoreTotal + (45 * NumberPyrovore)

MsgBox "You have " & NumberPyrovore & " worth " & PyrovoreTotal

End Sub

Private Sub cmdPyrovoreTotalCls_Click()
PyrovoreTotal = 0

End Sub

Private Sub Form_Load()

picPicture.AutoSize = True
picPicture.Picture = LoadPicture(App.Path & "\pyrovore.JPG")

End Sub
