VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404080&
   Caption         =   "Login"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelloU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtHello 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtInAt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   19
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdInvInf 
      Caption         =   "INVENTORY INFORMATION"
      Height          =   735
      Left            =   3000
      TabIndex        =   18
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdFinInf 
      Caption         =   "FINANCIAL INFO"
      Height          =   735
      Left            =   4440
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmpInf 
      Caption         =   "EMPLOYEE INFORMATION"
      Height          =   735
      Left            =   3000
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSchInf 
      Caption         =   "SCHEDULE INFORMATION"
      Height          =   735
      Left            =   1560
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtClkIn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   3480
      TabIndex        =   13
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5160
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   2640
   End
   Begin VB.CommandButton cmdJobInf 
      Caption         =   "JOB INFORMATION"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmpTask 
      Caption         =   "EMPLOYEE TASK"
      Height          =   735
      Left            =   1560
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmpSch 
      Caption         =   "EMPLOYEE SCHEDULE"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdClkOut 
      Caption         =   "CLOCK OUT"
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdClkIn 
      Caption         =   "CLOCK IN"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      Caption         =   "WELCOME TO ABBEY WOODWORKING"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      Caption         =   "Input Password"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      Caption         =   "Input Username"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    cmdEmpSch.Visible = False
    cmdJobInf.Visible = False
    cmdEmpTask.Visible = False
    cmdClkIn.Visible = False
    cmdClkOut.Visible = False
    cmdEmpInf.Visible = False
    cmdSchInf.Visible = False
    cmdInvInf.Visible = False
    cmdFinInf.Visible = False
    txtTime.Visible = False
    txtDate.Visible = False
    txtHello.Visible = False
    txtClkIn.Visible = False
    cmdLogout.Visible = False
    txtHelloU.Visible = False
    
    
End Sub

Private Sub cmdLogin_Click()

Dim FName As String
Dim Password As String
Dim Found As Boolean
Dim RecN As Integer
Dim DB As Database
Dim RS As Recordset2


Set DB = OpenDatabase(App.Path & "\WoodShop.accdb")
Set RS = DB.OpenRecordset("EmployeeInfo")


FName = txtUser.Text
Password = txtPass.Text
Found = False
RecN = 0
txtHello.Text = "Hello"



    Do Until RS.EOF Or Found = True
        RecN = RecN + 1
        If LCase(RS![FirstName]) = LCase(FName) And LCase(RS![Password]) = LCase(Password) Then
            Found = True
            txtHelloU.Text = RS![FirstName]
            If (RS![EmployedAs]) = "Employee" Then
                cmdClkIn.Visible = True
                cmdClkOut.Visible = True
                cmdEmpSch.Visible = True
                cmdEmpTask.Visible = True
                txtTime.Visible = True
                txtDate.Visible = True
                txtHello.Visible = True
                txtClkIn.Visible = True
                cmdLogout.Visible = True
                txtHelloU.Visible = True
                        
            ElseIf (RS![EmployedAs]) = "Manager" Then
                cmdClkIn.Visible = True
                cmdClkOut.Visible = True
                txtTime.Visible = True
                txtDate.Visible = True
                cmdJobInf.Visible = True
                cmdEmpTask.Visible = True
                txtHello.Visible = True
                txtClkIn.Visible = True
                cmdLogout.Visible = True
                cmdEmpInf.Visible = True
                cmdSchInf.Visible = True
                cmdInvInf.Visible = True
                cmdFinInf.Visible = True
                txtHelloU.Visible = True
                
            End If
            
        End If
        RS.MoveNext
    Loop

End Sub
Private Sub cmdClkIn_Click()
    
    MsgBox "YOU ARE NOW CLOCKED IN"
    txtInAt.Text = "CLOCKED IN AT"
    txtClkIn.Text = txtTime.Text
    
End Sub
Private Sub cmdClkOut_Click()
    
    MsgBox ("YOU ARE NOW CLOCKED OUT AT " & txtTime.Text)
    txtInAt.Text = ""
    txtClkIn.Text = ""
End Sub
Private Sub cmdEmpSch_Click()
    frmEmployeeSch.Show
    frmLogin.Hide
End Sub

Private Sub cmdLogout_Click()
    
    cmdEmpSch.Visible = False
    cmdJobInf.Visible = False
    cmdEmpTask.Visible = False
    cmdClkIn.Visible = False
    cmdClkOut.Visible = False
    cmdEmpInf.Visible = False
    cmdSchInf.Visible = False
    cmdInvInf.Visible = False
    cmdFinInf.Visible = False
    txtTime.Visible = False
    txtDate.Visible = False
    txtHello.Visible = False
    txtClkIn.Visible = False
    cmdLogout.Visible = False
    txtHelloU.Visible = False
    
    txtUser.Text = ""
    txtPass.Text = ""
    
End Sub


Private Sub Timer1_Timer()
    txtTime.Text = Format(Now, "hh:mm:ss")
End Sub

Private Sub Timer2_Timer()
    txtDate.Text = Format(Now, "mm/dd/yyyy")
End Sub

