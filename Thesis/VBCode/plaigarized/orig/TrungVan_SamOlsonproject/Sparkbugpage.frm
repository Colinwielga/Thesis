VERSION 5.00
Begin VB.Form Sparkbugpage 
   Caption         =   "Find out our spark bug products"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form3"
   ScaleHeight     =   5280
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Look up"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox Picresult2 
      Height          =   375
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox partnumber 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox brandname 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox picresult 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "This is the part number of our spark plug that you are looking for"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Please enter the part number"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter the brand name (A,B,Z)"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   $"Sparkbugpage.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Sparkbugpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr2 As Integer, brand(1 To 10) As String, brandA(1 To 10) As String
Dim brandB(1 To 10) As String, brandZ(1 To 10) As String, found As Boolean



Private Sub Form_Activate()
 
  Open App.Path & "\brands.txt" For Input As #2
   ctr2 = 0
   Do While Not EOF(2)
      ctr2 = ctr2 + 1
      Input #2, brand(ctr2), brandA(ctr2), brandB(ctr2), brandZ(ctr2)
   Loop
   
   Close #2
        
End Sub

Private Sub Command1_Click()
   Dim j As Integer, b As String
   b = brandname
   
   picresult.Cls
      
   found = False
   
   For j = 1 To ctr2
       If Left(brandA(j), 1) = UCase(b) Then
          found = True
          picresult.Print brandA(j)
       ElseIf Left(brandB(j), 1) = UCase(b) Then
          found = True
          picresult.Print brandB(j)
       ElseIf Left(brandZ(j), 1) = UCase(b) Then
          found = True
          picresult.Print brandZ(j)
       End If
    Next j
    
    If (Not found) Then
        MsgBox ("We cannot find any information about this brand! Please enter another brand name to consult!")
    Else
        picresult.Print ""
        picresult.Print "This is all the spark plug product of brand "; UCase(b)
    End If
   
End Sub

Private Sub Command2_Click()
   Dim k As Integer, p As String
   
   p = partnumber
   
   k = 0
   found = False
   
   Do While ((Not found) And (k < ctr2))
      k = k + 1
      If brandA(k) = UCase(p) Then
         found = True
      ElseIf brandB(k) = UCase(p) Then
         found = True
      ElseIf brandZ(k) = UCase(p) Then
         found = True
      End If
   Loop
   
   If (Not found) Then
       MsgBox ("We do not have a spark plug compatible to this one at the moment. Please try again later!")
   Else
       Picresult2.Print brand(k)
   End If
      
End Sub
Private Sub Command3_Click()
   generalpage.Show
   Sparkbugpage.Hide
End Sub
