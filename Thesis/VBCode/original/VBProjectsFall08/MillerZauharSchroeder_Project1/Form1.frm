VERSION 5.00
Begin VB.Form frmBegin 
   BackColor       =   &H00800000&
   Caption         =   "FormBegin"
   ClientHeight    =   12795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19485
   LinkTopic       =   "Form1"
   ScaleHeight     =   12795
   ScaleWidth      =   19485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6735
      Left            =   1320
      ScaleHeight     =   6735
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton cmdrules 
      Caption         =   "Click Here to See the Rules"
      Height          =   975
      Left            =   7920
      TabIndex        =   2
      Top             =   5520
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   6240
      ScaleHeight     =   4155
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Do You Want a Million Dollars??? Click Here to Start"
      Height          =   975
      Left            =   7080
      TabIndex        =   0
      Top             =   6600
      Width           =   4215
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdrules_Click()
Dim ruleslines(1 To 50) As String
Dim counter As Integer
Dim n As Integer
PicResults.Cls


Open App.Path & "\rules.txt" For Input As #1
    Do Until EOF(1)
    counter = counter + 1
    Input #1, ruleslines(counter)
    Loop
Close #1

For n = 1 To counter
PicResults.Print ruleslines(n)
Next n




End Sub

Private Sub cmdStart_Click()
FiftyEnabled = True
PhoneEnabled = True
AudienceEnabled = True

frmBegin.Hide
frmCharacters.Show
End Sub

Private Sub Form_Load()
'FiftyEnabled = True
'PhoneEnabled = True
'AudienceEnabled = True

End Sub

