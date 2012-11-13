VERSION 5.00
Begin VB.Form frmtokyo 
   Caption         =   "Tokyo"
   ClientHeight    =   6930
   ClientLeft      =   3870
   ClientTop       =   1860
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Display"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   12
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdnation 
      Caption         =   "See the Selections"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdrank 
      Caption         =   "See the Top 10"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdanswerjp 
      Caption         =   "Get Resut"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdgdpjp 
      Caption         =   "What is GDP??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtgdpjp 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton Cmdquitjp 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H80000018&
      Height          =   4455
      Left            =   4920
      ScaleHeight     =   4395
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label lblnation 
      BackColor       =   &H80000002&
      Caption         =   "Which is Chinese, Korean and Japanese??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Label lblgdpjp 
      BackColor       =   &H80000002&
      Caption         =   "How much is the recent GDP of Japan??    eg) US: 11trillion 750billion"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "TOKYO, JAPAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   8070
      Left            =   0
      Picture         =   "frmtokyo.frx":0000
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "frmtokyo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nation(1 To 10) As String
Dim X As Single
Dim gdp(1 To 10) As Single


Private Sub cmdquit_Click()

End


End Sub

Private Sub cmdanswerjp_Click()
'This button determine how user's estimate of Japan's
'GDP is close to the each country's GDP.
Dim answer As Single

answer = txtgdpjp.Text
X = 1
Do Until answer > gdp(X)
X = X + 1
Loop

picoutput.Print gdp(X), nation(X)


End Sub

Private Sub cmdclear_Click()
picoutput.Cls

End Sub

Private Sub cmdgdpjp_Click()
'Explain about GDP by message box.
MsgBox "GDP: Gross Domestic Production...GDP is the total value of goods and services produced by a nation.", , "GDP"
End Sub

Private Sub cmdload_Click()

'This button is for loadning the GDP data from the file.

Open App.Path & "\gdp.txt" For Input As #1
For X = 1 To 10
    Input #1, gdp(X), nation(X)
Next X


End Sub

Private Sub cmdmain_Click()
frmtokyo.Hide
frmmain.Show
End Sub

Private Sub cmdnation_Click()
frmtokyo.Hide
frmnation.Show

End Sub

Private Sub Cmdquitjp_Click()
End

End Sub

Private Sub Label2_Click()

End Sub

Private Sub cmdrank_Click()

'This conde shows the top ten of highest GDP in the world.


For X = 1 To 10
picoutput.Print gdp(X), nation(X)
Next X
End Sub

