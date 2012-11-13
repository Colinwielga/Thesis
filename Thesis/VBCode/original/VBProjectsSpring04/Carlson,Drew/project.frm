VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Welcome to the SJU Hockey Statistics Program"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Picture         =   "project.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdassists 
      Caption         =   "Click to see team assists leaders"
      Enabled         =   0   'False
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
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdgoals 
      Caption         =   "Click to see team goal leaders"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to enter a player's number and view his stats"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H000000FF&
      FillColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7095
      Left            =   3480
      Picture         =   "project.frx":16DC
      ScaleHeight     =   7035
      ScaleWidth      =   7995
      TabIndex        =   5
      Top             =   0
      Width           =   8055
   End
   Begin VB.CommandButton cmdloadarray 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to Load Statistics into a file "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Click to end the SJU Hockey Statistics Program"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdleaders 
      Caption         =   "Click to see team point leaders"
      Enabled         =   0   'False
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
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Click to search roster for a name"
      Enabled         =   0   'False
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
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdteampic 
      Caption         =   "Click to see Saint John's Hockey team"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by: Drew Carlson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SJU Hockey Statistics Program (SJUhockeyprogram.vbp)
'Form1 (project.frm)
'Drew Carlson
'3/14/04
' This form acts as the main page. It offers the user options as to what he/she can do with the program.
' This project is a program which takes offensive stats of the SJU hockey team for 2003-2004 season
' and sorts the data multiple ways. The program also displays pictures of the team and allows the
' user to search the data.
Option Explicit
'Assigning the variables data types.
Dim path As String
Dim num(1 To 27) As Integer
Dim names(1 To 27) As String
Dim gp(1 To 27) As Integer
Dim goals(1 To 27) As Integer
Dim assists(1 To 27) As Integer
Dim points(1 To 27) As Integer
Dim player As Integer
Dim found As Boolean
Dim CTR As Integer
Dim Temp, Temp1, Temp2, Temp3, Temp4, Temp5 As Integer
Dim Q As Integer, a As Integer
Dim P As Integer
Dim I As Integer

'This button sorts the team assist leaders from greatest to least
'and print the list out.

Private Sub cmdassists_Click()
Dim C, P, I As Integer
picresults.Cls
For P = 1 To 26
    For C = 1 To 27 - P
        If assists(C) < assists(C + 1) Then
            Temp = assists(C)
            assists(C) = assists(C + 1)
            assists(C + 1) = Temp
            
            Temp1 = num(C)
            num(C) = num(C + 1)
            num(C + 1) = Temp1
            
            Temp2 = names(C)
            names(C) = names(C + 1)
            names(C + 1) = Temp2
            
            Temp3 = gp(C)
            gp(C) = gp(C + 1)
            gp(C + 1) = Temp3
            
            Temp4 = points(C)
            points(C) = points(C + 1)
            points(C + 1) = Temp4
            
            Temp5 = goals(C)
            goals(C) = goals(C + 1)
            goals(C + 1) = Temp5
        End If
    Next C
Next P
picresults.Print "#"; Tab(10); "Names"; Tab(33); "GP"; Tab(39); "G"; Tab(44); "A"; Tab(49); "P"
picresults.Print "***************************************************************************"
For I = 1 To 24
    picresults.Print ; num(I); Tab(10); names(I); Tab(33); gp(I); Tab(39); goals(I); Tab(44); assists(I); Tab(49); points(I)
Next I
End Sub
'This button ends the program
Private Sub cmdend_Click()
MsgBox ("Thank you for using the SJU Hockey Statistics Program.")
End
End Sub

'This button sorts the team goal scorers from least to greatest and
'prints the list out.

Private Sub cmdgoals_Click()
Dim C, P, I As Integer
picresults.Cls
For P = 1 To 26
    For C = 1 To 27 - P
        If goals(C) < goals(C + 1) Then
            Temp = goals(C)
            goals(C) = goals(C + 1)
            goals(C + 1) = Temp
            
            Temp1 = num(C)
            num(C) = num(C + 1)
            num(C + 1) = Temp1
            
            Temp2 = names(C)
            names(C) = names(C + 1)
            names(C + 1) = Temp2
            
            Temp3 = gp(C)
            gp(C) = gp(C + 1)
            gp(C + 1) = Temp3
            
            Temp4 = points(C)
            points(C) = points(C + 1)
            points(C + 1) = Temp4
            
            Temp5 = assists(C)
            assists(C) = assists(C + 1)
            assists(C + 1) = Temp5
        End If
    Next C
Next P
picresults.Print "#"; Tab(10); "Names"; Tab(33); "GP"; Tab(39); "G"; Tab(44); "A"; Tab(49); "P"
picresults.Print "***************************************************************************"
For I = 1 To 24
    picresults.Print ; num(I); Tab(10); names(I); Tab(33); gp(I); Tab(39); goals(I); Tab(44); assists(I); Tab(49); points(I)
Next I
End Sub

' This button sorts the team point leaders from greatest to least
'and prints the list out.

Private Sub cmdleaders_Click()
Dim C, P, I As Integer
picresults.Cls
For P = 1 To 26
    For C = 1 To 27 - P
        If points(C) < points(C + 1) Then
            Temp = points(C)
            points(C) = points(C + 1)
            points(C + 1) = Temp
            
            Temp1 = num(C)
            num(C) = num(C + 1)
            num(C + 1) = Temp1
            
            Temp2 = names(C)
            names(C) = names(C + 1)
            names(C + 1) = Temp2
            
            Temp3 = gp(C)
            gp(C) = gp(C + 1)
            gp(C + 1) = Temp3
            
            Temp4 = goals(C)
            goals(C) = goals(C + 1)
            goals(C + 1) = Temp4
            
            Temp5 = assists(C)
            assists(C) = assists(C + 1)
            assists(C + 1) = Temp5
        End If
    Next C
Next P
picresults.Print "#"; Tab(10); "Names"; Tab(33); "GP"; Tab(39); "G"; Tab(44); "A"; Tab(49); "P"
picresults.Print "***************************************************************************"
For I = 1 To 24
    picresults.Print ; num(I); Tab(10); names(I); Tab(33); gp(I); Tab(39); goals(I); Tab(44); assists(I); Tab(49); points(I)
Next I
    End Sub
    
'This button loads the data in a text file into arrays so the data
' is saved for later use of other commands.

Private Sub cmdloadarray_Click()
Q = 0
'Open ("M:\Carlson_Drew\NEWstatssheet.txt") For Input As #1
Open "N:\CS130\handin\Carlson, Drew\NEWstatssheet.txt" For Input As #1
Do While Not EOF(1)
    Q = Q + 1
    Input #1, num(Q), names(Q), gp(Q), goals(Q), assists(Q), points(Q)
Loop


Close (1)
cmdassists.Enabled = True
cmdgoals.Enabled = True
cmdleaders.Enabled = True
cmdsearch.Enabled = True
cmdview.Enabled = True
End Sub

'This button allows the user to enter a first and last name to see
' if the name entered is listed on the roster as an active player.

Private Sub cmdsearch_Click()
Dim found As Boolean
found = False
Dim a As String
'Ask the user for a player to search for
a = InputBox("Enter the first and last name you wish to search the roster for. [Capslock Sensitive]")
I = 1
picresults.Cls
Do While I <= 27 And found = False
    If a = names(I) Then found = True
    I = I + 1
Loop
If found = True And a <> "" Then
picresults.Print a; " is listed on the roster as an active player."
Else: picresults.Print a; " is not listed on the roster as an active player."
End If
End Sub

'This button shows the user a team picture of SJU hockey

Private Sub cmdteampic_Click()
Form2.Show
Form1.Hide

End Sub

'This button allows the user to look at an isolated players' stats given that the user enters his jersey number.

Private Sub cmdview_Click()
Dim found As Boolean
picresults.Cls
'Ask the user for a jersey number to search for
a = InputBox("Enter a player's number.")
found = False
P = 1
Do While P <= 27 And found = False
    If a = num(P) Then found = True
    P = P + 1
Loop

If found = True Then
picresults.Print "#"; Tab(10); "Name"; Tab(33); "GP"; Tab(39); "G"; Tab(44); "A"; Tab(49); "P"
picresults.Print "***************************************************************************"
picresults.Print ; num(P - 1); Tab(10); names(P - 1); Tab(33); gp(P - 1); Tab(39); goals(P - 1); Tab(44); assists(P - 1); Tab(49); points(P - 1)
Else: picresults.Print "Number"; a; "is not on the roster."
End If
End Sub

