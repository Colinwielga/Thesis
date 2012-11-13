VERSION 5.00
Begin VB.Form frmCustomerChoose 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Name Selection"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00008000&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      MaskColor       =   &H80000007&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdJoinNow 
      BackColor       =   &H00008000&
      Caption         =   "Not a member? Sign up today!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   4095
   End
   Begin VB.ComboBox cboxMembers 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Text            =   "Please Select Your Name"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblmember 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Member Selection"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   6495
   End
End
Attribute VB_Name = "frmCustomerChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This bank system was designed and created by Mark Brown and David Bernardy

Option Explicit


Private Sub cmdJoinNow_Click()

frmCustomerChoose.Hide                          'Makes the form frmCustomerChoose hidden
frmSignUp.Show                                  'Makes the form frmSignUp visible

End Sub

Private Sub cmdNext_Click()

position = cboxMembers.ListIndex                'Sets the member selected from the drop down menu equal to a value for later use

frmCustomerChoose.Hide                         'Makes the form frmCustomerChoose hidden
frmCustomerIdentity.Show                        'Makes the form frmCustomerIdentity visible

frmCustomerIdentity.lblmember.Caption = firstname(position + 1) & " " & lastname(position + 1)      'Makes the label box on the formCustomerIdentity show the member's name
frmCustomerIdentity.picIdentity.Picture = LoadPicture(App.Path & "\" & id(position + 1))            'Makes the picture box on the form frmCustomerIdentity show the member's picture

End Sub

Private Sub cmdQuit_Click()
End                                         'Exits the bank

End Sub

Private Sub Form_Load()
Dim I As Integer
ReDim lastname(1 To 500) As String
ReDim firstname(1 To 500) As String
ReDim accountnum(1 To 500) As Long
ReDim streetadd(1 To 500) As String
ReDim city(1 To 500) As String
ReDim state(1 To 500) As String
ReDim zipcode(1 To 500) As Long
ReDim password(1 To 500) As String
ReDim checkingbal(1 To 500) As Double
ReDim savingsbal(1 To 500) As Double
ReDim id(1 To 500) As String


Open App.Path & "\Members.txt" For Input As #1      'Opens the data file

ctr = 0

Do Until EOF(1)                                     'fills the array with ctr numbers from the data file
    ctr = ctr + 1
    Input #1, lastname(ctr), firstname(ctr), accountnum(ctr), streetadd(ctr), city(ctr), state(ctr), zipcode(ctr), password(ctr), savingsbal(ctr), checkingbal(ctr), id(ctr)
Loop

last = ctr

Close #1

Dim comp, pass As Integer
Dim temp1 As String
Dim temp2 As String
Dim temp3 As Long
Dim temp4 As String
Dim temp5 As String
Dim temp6 As String
Dim temp7 As Long
Dim temp8 As String
Dim temp9 As Double
Dim temp10 As Double
Dim temp11 As String

'sort by last name A to Z by a bubble sort
For pass = 1 To ctr - 1                                 'Keeps track of how many passes
    For comp = 1 To ctr - pass                          'Kepps track of how many comparisons
        If lastname(comp) > lastname(comp + 1) Then
            temp1 = lastname(comp)                      'Exchanges values if out of order
            lastname(comp) = lastname(comp + 1)
            lastname(comp + 1) = temp1
            
            temp2 = firstname(comp)                     'Exchanges corresponding member information
            firstname(comp) = firstname(comp + 1)
            firstname(comp + 1) = temp2
            
            temp3 = accountnum(comp)                     'Exchanges corresponding member information
            accountnum(comp) = accountnum(comp + 1)
            accountnum(comp + 1) = temp3
            
            temp4 = streetadd(comp)                      'Exchanges corresponding member information
            streetadd(comp) = streetadd(comp + 1)
            streetadd(comp + 1) = temp4
            
            temp5 = city(comp)                           'Exchanges corresponding member information
            city(comp) = city(comp + 1)
            city(comp + 1) = temp5
            
            temp6 = state(comp)                          'Exchanges corresponding member information
            state(comp) = state(comp + 1)
            state(comp + 1) = temp6
            
            temp7 = zipcode(comp)                        'Exchanges corresponding member information
            zipcode(comp) = zipcode(comp + 1)
            zipcode(comp + 1) = temp7
            
            temp8 = password(comp)                       'Exchanges corresponding member information
            password(comp) = password(comp + 1)
            password(comp + 1) = temp8
            
            temp9 = savingsbal(comp)                     'Exchanges corresponding member information
            savingsbal(comp) = savingsbal(comp + 1)
            savingsbal(comp + 1) = temp9
            
            temp10 = checkingbal(comp)                    'Exchanges corresponding member information
            checkingbal(comp) = checkingbal(comp + 1)
            checkingbal(comp + 1) = temp10
            
            temp11 = id(comp)                             'Exchanges corresponding member information
            id(comp) = id(comp + 1)
            id(comp + 1) = temp11
            
        End If
    Next comp
Next pass


'Displays members Last name, first name and account number in the dropdown box
For I = 1 To ctr
    Dim x As String
    x = lastname(I) & ", " & firstname(I) & "    " & accountnum(I)
    cboxMembers.AddItem x                                               'adds items to dropdown list
Next I


End Sub
