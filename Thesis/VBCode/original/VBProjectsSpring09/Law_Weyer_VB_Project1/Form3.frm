VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form3"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form3"
   ScaleHeight     =   7935
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   7080
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   240
      ScaleHeight     =   5655
      ScaleWidth      =   11775
      TabIndex        =   6
      Top             =   120
      Width           =   11775
   End
   Begin VB.CommandButton cmdWartburg 
      Caption         =   "Wartburg"
      Height          =   615
      Left            =   10320
      TabIndex        =   5
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdVenice 
      Caption         =   "Venice"
      Height          =   615
      Left            =   8160
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdRome 
      Caption         =   "Rome"
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrague 
      Caption         =   "Prague"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBudapest 
      Caption         =   "Budapest"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBerlin 
      Caption         =   "Berlin"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Form3.Hide      'All these pictures were taken by Randy or used with permission from Alec Losinski, a friend of Randy's
Form1.Show      'Also to note on this form:there is only one counter for the whole form. We tried doing a counter for each set of pictures
End Sub         'But found out that it was unnessecary, since the counter gets reset after the last picture.
                
Private Sub cmdBerlin_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Berlin.JPG")
    Else
    picResults.Picture = LoadPicture(App.Path & "\Berlin.JPG")
    PicCtr = 0
    End If
End Sub

Private Sub cmdBudapest_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Budapest.JPG")
    ElseIf PicCtr = 2 Then
    picResults.Picture = LoadPicture(App.Path & "\Budapest1.JPG")
       PicCtr = 0
End If
End Sub

Private Sub cmdPrague_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Prague.JPG")
    Else
    picResults.Picture = LoadPicture(App.Path & "\Prague.JPG")
    PicCtr = 0
    End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub


Private Sub cmdRome_Click()

PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Rome1.JPG")
    ElseIf PicCtr = 2 Then
    picResults.Picture = LoadPicture(App.Path & "\Rome2.JPG")
    ElseIf PicCtr = 3 Then
    picResults.Picture = LoadPicture(App.Path & "\Rome3.JPG")
    PicCtr = 0
End If

End Sub

Private Sub cmdSalzburg_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Salzburg.JPG")
    ElseIf PicCtr = 2 Then
    picResults.Picture = LoadPicture(App.Path & "\Salzburg2.JPG")
       PicCtr = 0
End If
End Sub

Private Sub cmdVenice_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\Venice.JPG")
    ElseIf PicCtr = 2 Then
    picResults.Picture = LoadPicture(App.Path & "\Venice2.JPG")
    ElseIf PicCtr = 3 Then
    picResults.Picture = LoadPicture(App.Path & "\Venice3.JPG")
    PicCtr = 0
End If
End Sub

Private Sub cmdwartburg_Click()
PicCtr = PicCtr + 1
    If PicCtr = 1 Then
    picResults.Picture = LoadPicture(App.Path & "\wartburg1.JPG")
    Else
    picResults.Picture = LoadPicture(App.Path & "\wartburg1.JPG")
    PicCtr = 0
    End If
End Sub

