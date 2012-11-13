VERSION 5.00
Begin VB.Form frmTemp 
   BackColor       =   &H000080FF&
   Caption         =   "Temper"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Back to previous screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   5
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox picout 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   3720
      ScaleHeight     =   4875
      ScaleWidth      =   8595
      TabIndex        =   3
      Top             =   2520
      Width           =   8655
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtkids 
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblKids 
      Caption         =   "Do you have kids?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dogs(VB-project.vbp)
'Form Name: frmTemp (frmTemp.frm)
'Author: Libby Owen
'Date: Thursday Oct. 27
'Purpose: This form was written so that the user can input if they have kids or not
        ' and then be able to see a list of dogs that would work if they do or if they
        'do not have children. The user will get to see what kind of personality
        ' some dogs have.



Option Explicit
Dim x As String
Dim breeds(1 To 11) As String
Dim temp(1 To 11) As String
Dim I As Integer
Dim num(1 To 11) As Integer

Private Sub cmdList_Click()
'this button opens parallel arrays and lets the user look at some different breeds and what
'kind of temperaments the breed has
Open App.Path & "\temp.txt" For Input As #2
picout.Cls

For I = 1 To 11
    Input #2, breeds(I), temp(I), num(I)
    picout.Print breeds(I); Tab(25); temp(I); Tab(60)
Next I
Close #2



End Sub


Private Sub cmdGo_Click()
'this button is what the user will click after they input if they have kids or not
'it will tell them what type of dog that will work for them

x = txtkids.Text
picout.Cls

If x = "yes" Then
    picout.Print "Since you have kids you will need a dog with an even "
    picout.Print ; "temperament. Some examples of these breeds are...  "
    picout.Print ; "                                "
    Open App.Path & "\temp.txt" For Input As #2
    
    For I = 1 To 11
        Input #2, breeds(I), temp(I), num(I)
        
        If num(I) = 0 Then
            picout.Print breeds(I)
        End If
    Next I
    
ElseIf x = "no" Then
    picout.Print "You don't have kids so the temperament "
    picout.Print ; " depends on what you want."
    picout.Print ; "                                  "
    picout.Print ; "Click on the list button to see a list of"
    picout.Print ; " breeds and their temperaments"




    
End If







End Sub



Private Sub cmdQuit_Click()
frmTemp.Hide
frmFind.Show
End Sub
