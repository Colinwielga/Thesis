VERSION 5.00
Begin VB.Form frmEmployeeSch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Employee Schedule"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFName 
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picFri 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox picThur 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picWed 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picTue 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox picMon 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   855
      Left            =   5160
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "VIEW SCHEDULE"
      Height          =   1935
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtWeek 
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtMonth 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK TO MAIN"
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employees First Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "FRIDAY"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "THURSDAY"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "WEDNESDAY"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "TUESDAY"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "MONDAY"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Type Week Number (1 - 4)"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Type Month e.g. ""January"""
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmEmployeeSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmLogin.Show
    frmEmployeeSch.Hide
End Sub


Private Sub cmdView_Click()

Dim Month As String
Dim Week As Integer
Dim DB As Database
Dim RS As Recordset2
Dim RecN As Integer
Dim Found As Boolean
Dim FName As String

Month = LCase(txtMonth.Text)
FName = LCase(txtFName.Text)

Set DB = OpenDatabase(App.Path & "\JanScd.accdb")

    If Month = "january" Then
    Set RS = DB.OpenRecordset("January")
    Found = False
    RecN = 0
        Do Until RS.EOF Or Found = True
            RecN = RecN + 1
            If LCase(RS![FName]) = LCase(FName) And RS![Week] = txtWeek.Text Then
                Found = True
                picMon.Print "Hours are " & RS![Monday]
                picTue.Print "Hours are " & RS![Tuesday]
                picWed.Print "Hours are " & RS![Wednesday]
                picThur.Print "Hours are " & RS![Thursday]
                picFri.Print "Hours are " & RS![Friday]
            End If
            RS.MoveNext
        Loop
    ElseIf Month = "february" Then
    Set RS = DB.OpenRecordset("February")
    Found = False
    RecN = 0
        Do Until RS.EOF Or Found = True
            RecN = RecN + 1
            If LCase(RS![FName]) = LCase(FName) And RS![Week] = txtWeek.Text Then
                Found = True
                picMon.Print "Hours are " & RS![Monday]
                picTue.Print "Hours are " & RS![Tuesday]
                picWed.Print "Hours are " & RS![Wednesday]
                picThur.Print "Hours are " & RS![Thursday]
                picFri.Print "Hours are " & RS![Friday]
            End If
            RS.MoveNext
        Loop
    ElseIf Month = "march" Then
    Set RS = DB.OpenRecordset("March")
    Found = False
    RecN = 0
        Do Until RS.EOF Or Found = True
            RecN = RecN + 1
            If LCase(RS![FName]) = LCase(FName) And RS![Week] = txtWeek.Text Then
                Found = True
                picMon.Print "Hours are " & RS![Monday]
                picTue.Print "Hours are " & RS![Tuesday]
                picWed.Print "Hours are " & RS![Wednesday]
                picThur.Print "Hours are " & RS![Thursday]
                picFri.Print "Hours are " & RS![Friday]
            End If
            RS.MoveNext
        Loop
    End If
    
End Sub
Private Sub cmdClear_Click()
    txtMonth.Text = ""
    txtWeek.Text = ""
    picMon.Cls
    picTue.Cls
    picWed.Cls
    picThur.Cls
    picFri.Cls
End Sub

