VERSION 5.00
Begin VB.Form Plans 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form4"
   Picture         =   "Plans.frx":0000
   ScaleHeight     =   11355
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Seats 
      Caption         =   "Choose Your Seats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sort by Price"
      Height          =   495
      Left            =   3840
      TabIndex        =   23
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sort by Name"
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Shows"
      Height          =   495
      Left            =   3840
      TabIndex        =   21
      Top             =   5040
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   6120
      ScaleHeight     =   6435
      ScaleWidth      =   5955
      TabIndex        =   20
      Top             =   4680
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Return"
      Height          =   495
      Left            =   11040
      MaskColor       =   &H0000C0C0&
      TabIndex        =   19
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   4215
      Left            =   120
      Picture         =   "Plans.frx":240042
      ScaleHeight     =   4155
      ScaleWidth      =   5835
      TabIndex        =   18
      Top             =   6960
      Width           =   5895
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear Order"
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton SchedDisplay 
      Caption         =   "Summary"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton E4 
      Caption         =   "4th"
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton E3 
      Caption         =   "3rd"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton E2 
      Caption         =   "2nd"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton E1 
      Caption         =   "1st"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   3615
      Left            =   6360
      ScaleHeight     =   3555
      ScaleWidth      =   5715
      TabIndex        =   6
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtEvent 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtStart 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtPrice 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Excel Center"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   24
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label asd 
      BackColor       =   &H0000C0C0&
      Caption         =   "March Events"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Press Each Number for a Ticket Purchase:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Order Tickets"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Your Order:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Order Information"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Event Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Start Time:       (ex:12:30pm)"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Price"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "Plans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St.PaulEvents
'Plans
'David Wutchiett
'February 24, 2010
'Enables user to access and sort a list of Excel Center Events.  Allows ticket purchase through textbox input. Totals up purchase.


Dim I As Integer
Dim Ctr As Integer
Dim ev(1 To 31) As String
Dim da(1 To 31) As Integer
Dim pr(1 To 31) As Single
Dim ti(1 To 31) As String


Private Sub Clear_Click()
picResults.Cls
End Sub

Private Sub Command1_Click()
Form1.Show
Plans.Hide

End Sub

Private Sub Command2_Click()



Picture1.Print "March"; Tab(12); "Performance"; Tab(42); "Price"; Tab(55); "Time"

Ctr = 0
  
  
Open App.Path & "\event.txt" For Input As #1
For I = 1 To 31
        Input #1, da(I), ev(I), pr(I), ti(I)
        Ctr = Ctr + 1
        Picture1.Print da(I); Tab(12); ev(I); Tab(42); FormatCurrency(pr(I)); Tab(55); ti(I)
Next I
        

End Sub

Private Sub Command3_Click()

Picture1.Cls

Picture1.Print "March"; Tab(12); "Performance"; Tab(42); "Price"; Tab(55); "Time"

Dim pass, comp, J As Integer
Dim temp As String
Dim Pos As Integer

For pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - pass
        If ev(Pos) > ev(Pos + 1) Then
            temp = ev(Pos)
            ev(Pos) = ev(Pos + 1)
            ev(Pos + 1) = temp
            
            temp = pr(Pos)
            pr(Pos) = pr(Pos + 1)
            pr(Pos + 1) = temp
            
            temp = da(Pos)
            da(Pos) = da(Pos + 1)
            da(Pos + 1) = temp
            
            temp = ti(Pos)
            ti(Pos) = ti(Pos + 1)
            ti(Pos + 1) = temp
            
        End If
    Next Pos
Next pass

For J = 1 To 31
    Picture1.Print da(J); Tab(12); ev(J); Tab(42); FormatCurrency(pr(J)); Tab(55); ti(J)
Next J
End Sub

Private Sub Command4_Click()

Picture1.Cls

Picture1.Print "March"; Tab(12); "Performance"; Tab(42); "Price"; Tab(55); "Time"

Dim pass, comp, J As Integer
Dim temp As String
Dim Pos As Integer

For pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - pass
        If pr(Pos) > pr(Pos + 1) Then
            temp = pr(Pos)
            pr(Pos) = pr(Pos + 1)
            pr(Pos + 1) = temp
            
            temp = ev(Pos)
            ev(Pos) = ev(Pos + 1)
            ev(Pos + 1) = temp
            
            temp = da(Pos)
            da(Pos) = da(Pos + 1)
            da(Pos + 1) = temp
            
            temp = ti(Pos)
            ti(Pos) = ti(Pos + 1)
            ti(Pos + 1) = temp
            
        End If
    Next Pos
Next pass

For J = 1 To 31
    Picture1.Print da(J); Tab(12); ev(J); Tab(42); FormatCurrency(pr(J)); Tab(55); ti(J)
Next J

End Sub

Private Sub Command5_Click()

    Total = SunETime1 + SunETime2 + SunETime3 + SunETime4
MsgBox ("St.Paul Events does not add hidden taxes or fees!")
picResults.Print "***********************************************************************"
picResults.Print "Total:"; Tab(42); FormatCurrency(Total)

Seats.Visible = True
    
End Sub

Private Sub Command6_Click()

End Sub

Private Sub E1_Click()

SunOcca1 = txtEvent.Text
SunBTime1 = txtStart.Text
SunETime1 = txtPrice.Text


End Sub

Private Sub E2_Click()

SunOcca2 = txtEvent.Text
SunBTime2 = txtStart.Text
SunETime2 = txtPrice.Text


End Sub

Private Sub E3_Click()

SunOcca3 = txtEvent.Text
SunBTime3 = txtStart.Text
SunETime3 = txtPrice.Text


End Sub

Private Sub E4_Click()

SunOcca4 = txtEvent.Text
SunBTime4 = txtStart.Text
SunETime4 = txtPrice.Text


End Sub

Private Sub SchedDisplay_Click()

picResults.Cls
picResults.Print "Performance"; Tab(25); "Time"; Tab(42); "Price"
picResults.Print "***********************************************************************"
picResults.Print


    picResults.Print SunOcca1; Tab(25); SunBTime1; Tab(42); FormatCurrency(SunETime1);
    picResults.Print "  "

 
    picResults.Print SunOcca2; Tab(25); SunBTime2; Tab(42); FormatCurrency(SunETime2);
    picResults.Print "  "


    picResults.Print SunOcca3; Tab(25); SunBTime3; Tab(42); FormatCurrency(SunETime3);
    picResults.Print "  "
        

    picResults.Print SunOcca4; Tab(25); SunBTime4; Tab(42); FormatCurrency(SunETime4);
    picResults.Print " "
    
End Sub

Private Sub Slope1_Click()
    
End Sub

Private Sub Seats_Click()

Seating.Show
Plans.Hide

End Sub
