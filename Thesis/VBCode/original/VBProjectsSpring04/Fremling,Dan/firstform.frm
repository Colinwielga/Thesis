VERSION 5.00
Begin VB.Form Bowlform 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "firstform.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmore 
      Caption         =   "Learn More"
      Height          =   495
      Left            =   9240
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdearn 
      Caption         =   "Sort by Earnings"
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox picbio 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   7875
      TabIndex        =   9
      Top             =   6120
      Width           =   7935
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load and Print Data"
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox picmike 
      Height          =   2295
      Left            =   8040
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox pictommy 
      Height          =   2295
      Left            =   10080
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox picchris 
      Height          =   2295
      Left            =   5880
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox picpatrick 
      Height          =   2295
      Left            =   2040
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox picpete 
      Height          =   2295
      Left            =   10080
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox picmika 
      Height          =   2295
      Left            =   8040
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox picbrad 
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox picwalterray 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Patrick Allen"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Brad Angelo"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Chris Barnes "
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Michael Haugen Jr."
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Tommy Jones"
      Height          =   255
      Left            =   10080
      TabIndex        =   15
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Pete Weber"
      Height          =   255
      Left            =   10080
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mika Koivuniemi"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Walter Ray Williams Jr"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
End
Attribute VB_Name = "Bowlform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : PBA bowlers (Dan Fremling's VB Project.vbp)
'Form Name: Bowlers (Bowlform)'
'Author: Dan Fremling '
'Date Written: March 13, 2004'
'Purpose of Form: To let someone look at the top 8 bowlers of the PBA
                    ' and to let them see their stats and pictures
                    ' arrange them by averages
                    ' pick a certain bowler
                    ' and learn more about the bowler
                    
Option Explicit
Dim path As String
Dim bowler(1 To 10) As String
Dim bdays(1 To 10) As Date
Dim town(1 To 10) As String
Dim state(1 To 10) As String
Dim hand(1 To 10) As String
Dim earnings(1 To 10) As Double
Dim tempa As Double
Dim tempb As String
Dim tempc As String
Dim tempd As String
Dim tempe As String
Dim tempf As String
Dim x As Integer
Dim n As Integer
Dim pass As Integer






Private Sub cmdload_Click()
Open path & "bowlarray.txt" For Input As #1

picbio.Print "Name"; Tab(25); "City"; Tab(40); "State"; Tab(50); "Hand"; Tab(60); "Birthdate"; Tab(80); "Career Earnings"
        For x = 1 To 8
        Input #1, bowler(x), town(x), state(x), hand(x), bdays(x), earnings(x)
        'prints the bolwer's and their information from the array'
        picbio.Print bowler(x); Tab(25); town(x); Tab(40); state(x); Tab(50); hand(x); Tab(60); bdays(x); Tab(80); FormatCurrency(earnings(x))
       Next x
       
    
Close #1
End Sub

Private Sub cmdearn_Click()
picbio.Cls
    'clears any info in picture box'
n = 8

For pass = 1 To n - 1
    For x = 1 To n - pass
        'swap earnings of(x) and earnings(x+1)'
        If earnings(x) < earnings(x + 1) Then
            tempa = earnings(x)
        earnings(x) = earnings(x + 1)
        earnings(x + 1) = tempa
        'swap name of bowler(x) and bowler (x+1)
            tempb = bowler(x)
        bowler(x) = bowler(x + 1)
        bowler(x + 1) = tempb
        'swap name of town(x) and town (x+1)
            tempc = town(x)
        town(x) = town(x + 1)
        town(x + 1) = tempc
        'swap name of state(x) and state(x+1)
            tempd = state(x)
        state(x) = state(x + 1)
        state(x + 1) = tempd
        'swap name of hand(x) and hand(x+1)
         tempe = hand(x)
        hand(x) = hand(x + 1)
        hand(x + 1) = tempe
        'swap bday(x) and bday(x+1)
            tempf = bdays(x)
        bdays(x) = bdays(x + 1)
        bdays(x + 1) = tempf
        End If
    Next x
Next pass
picbio.Print "Name"; Tab(25); "City"; Tab(40); "State"; Tab(50); "Hand"; Tab(60); "Birthdate"; Tab(80); "Career Earnings"
For x = 1 To n
    'prints the results in rank of earnings in descending order'
    picbio.Print x; bowler(x); Tab(25); town(x); Tab(40); state(x); Tab(50); hand(x); Tab(60); bdays(x); Tab(80); FormatCurrency(earnings(x))
Next x
End Sub
    



Private Sub cmdmore_Click()
'Goes to Second form'
Bowlform.Hide
Bioform.Show
End Sub

Private Sub Form_Load()
path = "N:\CS130\handin\Fremling, Dan\"
'path for all files'
picwalterray = LoadPicture(path & "walterrayjr.jpg")
picbrad = LoadPicture(path & "bradangelo.jpg")
picmika = LoadPicture(path & "mikakoivuniemi.jpg")
picchris = LoadPicture(path & "chrisbarnes.jpg")
picmike = LoadPicture(path & "michaelhaugenjr.jpg")
picpatrick = LoadPicture(path & "patrickallen.jpg")
picpete = LoadPicture(path & "peteweber.jpg")
pictommy = LoadPicture(path & "tommyjones.jpg")
End Sub



