VERSION 5.00
Begin VB.Form frmContact 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Contacts for your sports!"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optFive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      Top             =   2640
      Width           =   1935
   End
   Begin VB.OptionButton optFour 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.OptionButton optThree 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton optTwo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton optOne 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Calculations!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click() ' returns to previous screen (frmsecond)
    frmContact.Hide
    frmSecond.Show
End Sub

Private Sub Form_Load() 'sets option buttons captions as top 5 ranked
    optOne.Caption = rankname(1)
    optTwo.Caption = rankname(2)
    optThree.Caption = rankname(3)
    optFour.Caption = rankname(4)
    optFive.Caption = rankname(5)
End Sub

Private Sub optFive_Click() 'sets this button as true and rest as false
    optTwo = False
    optThree = False
    optFour = False
    optOne = False
 Dim A(1 To 100) As String
   Dim pos As Integer
   frmContact.Cls
   frmContact.Print "here are your "; rankname(5); " Results"
    If InStr(rankname(5), "Track-") <> 0 Then
            Open App.Path & "\trackcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Soccer") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Basket") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
               frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Football") <> 0 Then
            Open App.Path & "\footballcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Hockey") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Walking") <> 0 Then
            Open App.Path & "\Walkcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Video") <> 0 Then
               Open App.Path & "\videocontact.txt" For Input As #1
                    Do Until EOF(1)
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          ElseIf InStr(rankname(5), "Danc") <> 0 Then
            Open App.Path & "\dancecontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(5), "Golf") <> 0 Then
               Open App.Path & "\golfcontact.txt" For Input As #1
               Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          Close #1
          ElseIf InStr(rankname(5), "Weight") <> 0 Then
            Open App.Path & "\weightcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(5), "Base") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        Else
        
            frmContact.Print "grrrr"
        End If
End Sub

Private Sub optFour_Click() 'sets this button as true and rest as false
    optOne = False
    optTwo = False
    optThree = False
    optFive = False
    Dim A(1 To 100) As String
    Dim pos As Integer
    frmContact.Cls
    frmContact.Print "here are your "; rankname(4); " Results"
     If InStr(rankname(4), "Track-") <> 0 Then
             Open App.Path & "\trackcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1)
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Soccer") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Basket") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
               frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Football") <> 0 Then
            Open App.Path & "\footballcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Hockey") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Walking") <> 0 Then
            Open App.Path & "\Walkcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Video") <> 0 Then
               Open App.Path & "\videocontact.txt" For Input As #1
                    Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          ElseIf InStr(rankname(4), "Danc") <> 0 Then
            Open App.Path & "\dancecontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(4), "Golf") <> 0 Then
               Open App.Path & "\golfcontact.txt" For Input As #1
               Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          Close #1
          ElseIf InStr(rankname(4), "Weight") <> 0 Then
            Open App.Path & "\weightcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(4), "Base") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        Else
        
            frmContact.Print "grrrr"
        End If
End Sub

Private Sub optOne_Click() 'sets this button as true and rest as false
    optTwo = False
    optThree = False
    optFour = False
    optFive = False
  Dim A(1 To 100) As String
   Dim pos As Integer
   frmContact.Cls
   frmContact.Print "here are your "; rankname(1); " Results"
    If InStr(rankname(1), "Track-") <> 0 Then
            Open App.Path & "\trackcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Soccer") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Basket") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
               frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Football") <> 0 Then
            Open App.Path & "\footballcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Hockey") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Walking") <> 0 Then
            Open App.Path & "\Walkcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Video") <> 0 Then
               Open App.Path & "\videocontact.txt" For Input As #1
                    Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\dancecontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(1), "Golf") <> 0 Then
               Open App.Path & "\golfcontact.txt" For Input As #1
               Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          Close #1
          ElseIf InStr(rankname(1), "Weight") <> 0 Then
            Open App.Path & "\weightcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(1), "Base") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        Else
        
            frmContact.Print "grrrr"
        End If
End Sub

Private Sub optThree_Click() 'sets this button as true and rest as false
    optOne = False
    optTwo = False
    optFour = False
    optFive = False

    Dim A(1 To 100) As String
    Dim pos As Integer
    frmContact.Cls
    frmContact.Print "here are your "; rankname(3); " Results"
    If InStr(rankname(3), "Track-") <> 0 Then
        Open App.Path & "\trackcontact.txt" For Input As #1
        pos = 0
        Do Until EOF(1) 'loads file to array
            pos = pos + 1
            Input #1, A(pos)
            frmContact.Print A(pos) ' displays results from file
        Loop
        Close #1
        ElseIf InStr(rankname(3), "Soccer") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Basket") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
               frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Football") <> 0 Then
            Open App.Path & "\footballcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Hockey") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Walking") <> 0 Then
            Open App.Path & "\Walkcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Video") <> 0 Then
               Open App.Path & "\videocontact.txt" For Input As #1
                    Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          ElseIf InStr(rankname(3), "Danc") <> 0 Then
            Open App.Path & "\dancecontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(3), "Golf") <> 0 Then
               Open App.Path & "\golfcontact.txt" For Input As #1
               Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          Close #1
          ElseIf InStr(rankname(3), "Weight") <> 0 Then
            Open App.Path & "\weightcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(3), "Base") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        Else
        
            frmContact.Print "grrrr"
        End If
End Sub

Private Sub optTwo_Click() 'sets this button as true and rest as false
    optOne = False
    optThree = False
    optFour = False
    optFive = False
   Dim A(1 To 100) As String
   Dim pos As Integer
   frmContact.Cls
   frmContact.Print "here are your "; rankname(2); " Results"
    If InStr(rankname(2), "Track-") <> 0 Then
            Open App.Path & "\trackcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Soccer") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Basket") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
               frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Football") <> 0 Then
            Open App.Path & "\footballcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Hockey") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Walking") <> 0 Then
            Open App.Path & "\Walkcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Video") <> 0 Then
               Open App.Path & "\videocontact.txt" For Input As #1
                    Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          ElseIf InStr(rankname(1), "Danc") <> 0 Then
            Open App.Path & "\dancecontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
          ElseIf InStr(rankname(2), "Golf") <> 0 Then
               Open App.Path & "\golfcontact.txt" For Input As #1
               Do Until EOF(1) 'loads file to array
                    pos = pos + 1
                    Input #1, A(pos)
                    frmContact.Print A(pos) ' displays results from file
                Loop
          Close #1
          ElseIf InStr(rankname(2), "Weight") <> 0 Then
            Open App.Path & "\weightcontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1) 'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        ElseIf InStr(rankname(2), "Base") <> 0 Then
            Open App.Path & "\misccontact.txt" For Input As #1
            pos = 0
            Do Until EOF(1)  'loads file to array
                pos = pos + 1
                Input #1, A(pos)
                frmContact.Print A(pos) ' displays results from file
            Loop
        Close #1
        Else
        
            frmContact.Print "grrrr"
        End If
End Sub
