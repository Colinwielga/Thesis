VERSION 5.00
Begin VB.Form frmpasodontuse 
   Caption         =   "El Paso"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      Caption         =   "Leave and continue on your tour de St. Joe"
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton cmdage 
      Caption         =   "Continue inside"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   1800
      Top             =   840
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Once inside the door...."
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "frmpasodontuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Project:  Tour De St. Joe
' Form:  frmpaso, "El Paso"
' author: Brooke and Josh
' Date written: 3/08/08
' Objective: To use if/then statments to figure out if someone could be allowed into this specific bar.

Private Sub cmdage_Click()
    
    Dim age As Integer
    Dim town As String
    

    age = InputBox("Can I see some I.D. please?  How old are you?", "Bouncer")
    

    If age >= 21 Then
        MsgBox "You're of age", , "OK"
            InputBox "Where are you from?"
                If town = "St. Joseph" Then
                    MsgBox "We've missed you!  Come in."
                ElseIf town <> "St. Joseph" Then
                    MsgBox "This is a townie bar - you should probably leave."
                End If
    ElseIf age < 21 Then
        MsgBox "Leave or I'm calling the police", , "Not a chance"
    End If


End Sub

Private Sub cmdleave_Click()
    
        frmpaso.Hide           'hides the matching form
        frmjoetown.Show        'shows the main form
    
End Sub
