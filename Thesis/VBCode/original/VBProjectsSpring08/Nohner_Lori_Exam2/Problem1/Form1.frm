VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdSnow 
      Caption         =   "Enter Snow Depth"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Depth As Single


Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSnow_Click()
Depth = InputBox("Enter the depth of the snow.", "Snow Depth")
    
        If Depth >= 0 And Depth < 0.5 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for skating", , "Skating"
        ElseIf Depth >= 0.5 And Depth <= 6 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for sledding", , "Sledding"
        ElseIf Depth > 6 And Depth < 14 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for X-Country Skiing", , "X-Country Skiing"
        ElseIf Depth >= 14 And Depth <= 20 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for down-hill skiing", , "Down-Hill Skiing"
        ElseIf Depth > 20 And Depth <= 40 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for snow-shoeing", , "Snow-Shoeing"
        ElseIf Depth > 40 Then
            MsgBox "Snow " & FormatNumber(Depth, 3) & " inches deep is good for sitting by the fireplace", , "fireplace"
                Else
                    MsgBox "There is no snow!", , "No snow!"
       End If
End Sub
