VERSION 5.00
Begin VB.Form Shandong_cuisine 
   Caption         =   "Form1"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "Shandong_cuisine.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpassword 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   5400
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      Height          =   2895
      Left            =   5160
      ScaleHeight     =   2835
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdintroduce 
      Caption         =   "someinformation about Shandong cuisine"
      Height          =   1935
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Shandong_cuisine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub cmdintroduce_Click()
Dim sichuan(1 To 100) As String
Dim ctr As Integer
Dim n As Integer
picResults.Cls
Open App.Path & "\sichuan.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, sichuan(ctr)
    Loop
    Close #1
For n = 1 To ctr
    picResults.ForeColor = RGB(0, 0, 0)
    picResults.Print sichuan(n)
Next n
End Sub


Private Sub txtpassword_KeyPress(KeyAscii As Integer)
Dim Char As String, n As Integer
   Char = Chr(KeyAscii)
   KeyAscii = Asc(UCase(Char))
   Select Case Char
   Case "1"
   txtpassword.Refresh
  picResults.Picture = LoadPicture(App.Path & "\tanghua.jpg")
    Case "56"
        MsgBox " a B?"
    Case "39"
    n = KeyAscii
        picResults.Print "you pressed the letter 'p' and n is "; n
    Case "4"
        Text2.SetFocus
        Text2.Text = Char
    Case Else
        
    End Select
End Sub
