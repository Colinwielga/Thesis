VERSION 5.00
Begin VB.Form frmOlympics 
   BackColor       =   &H00008000&
   Caption         =   "Olympic Race Results"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChampions 
      BackColor       =   &H000000FF&
      Caption         =   "List of Champions and Times"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmd2004 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 2004 in Athens"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd2000 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 2000 in Sydney"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmd1996 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1996 in Atlanta"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmd1992 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1992 in Barcelona"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmd1988 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1988 in Seoul"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmd1984 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1984 in Los Angeles"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmd1980 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1980 in Moscow"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd1976 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1976 in Montreal"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmd1972 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1972 in Munich"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmd1968 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1968 in Mexico City"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmd1964 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1964 in Tokyo"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   6975
      Left            =   4320
      ScaleHeight     =   6915
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Race Results Page"
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmd1960 
      BackColor       =   &H000000FF&
      Caption         =   "Results for 1960 in Rome"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "frmOlympics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
'this button pull up results from 1960 olympic games'
Private Sub cmd1960_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Rome(1 To 7) As Integer, Name(1 To 7) As String, Minute(1 To 7) As Integer, Seconds(1 To 7) As Single
Dim D As Integer, E As Integer, Rome2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Integer, Seconds2(1 To 8) As Single
Open App.Path & "\5k1960.txt" For Input As #1
Open App.Path & "\10k1960.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Rome Results"
For A = 1 To 7
    Input #1, Rome(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 7
If Seconds(D) > 10 Then
    picResults.Print Rome(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Rome(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Rome Results"
For B = 1 To 8
    Input #2, Rome2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Rome2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Rome2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1964 olympic games'
Private Sub cmd1964_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Tokyo(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Integer, Seconds(1 To 8) As Single
Dim E As Integer, D As Integer, Tokyo2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Integer, Seconds2(1 To 8) As Single
Open App.Path & "\5k1964.txt" For Input As #1
Open App.Path & "\10k1964.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Tokyo Results"
For A = 1 To 8
    Input #1, Tokyo(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Tokyo(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Tokyo(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Tokyo Results"
For B = 1 To 8
    Input #2, Tokyo2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Tokyo2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Tokyo2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two steps close the 2 data files'
End Sub

'this button pull up results from 1968 olympic games'
Private Sub cmd1968_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Mexico(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Mexico2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1968.txt" For Input As #1
Open App.Path & "\10k1968.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Mexico City Results"
For A = 1 To 8
    Input #1, Mexico(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Mexico(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Mexico(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Mexico City Results"
For B = 1 To 8
    Input #2, Mexico2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Mexico2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Mexico2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these 2 steps close the 2 data files'
End Sub

'this button pull up results from 1972 olympic games'
Private Sub cmd1972_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Munich(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim E As Integer, D As Integer, Munich2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1972.txt" For Input As #1
Open App.Path & "\10k1972.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Munich Results"
For A = 1 To 8
    Input #1, Munich(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Munich(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Munich(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Munich Results"
For B = 1 To 8
    Input #2, Munich2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Munich2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Munich2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1976 olympic games'
Private Sub cmd1976_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Montreal(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Montreal2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1976.txt" For Input As #1
Open App.Path & "\10k1976.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Montreal Results"
For A = 1 To 8
    Input #1, Montreal(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Montreal(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Montreal(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Montreal Results"
For B = 1 To 8
    Input #2, Montreal2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Montreal2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Montreal2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1980 olympic games'
Private Sub cmd1980_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Moscow(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Moscow2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1980.txt" For Input As #1
Open App.Path & "\10k1980.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Moscow Results"
For A = 1 To 8
    Input #1, Moscow(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Moscow(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Moscow(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Moscow Results"
For B = 1 To 8
    Input #2, Moscow2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Moscow2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Moscow2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1984 olympic games'
Private Sub cmd1984_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer, D As Integer, E As Integer
Dim LA(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim LA2(1 To 7) As Integer, Name2(1 To 7) As String, Minute2(1 To 7) As Single, Seconds2(1 To 7) As Single
Open App.Path & "\5k1984.txt" For Input As #1
Open App.Path & "\10k1984.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Los Angeles Results"
For A = 1 To 8
    Input #1, LA(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print LA(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print LA(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Los Angeles Results"
For B = 1 To 7
    Input #2, LA2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 7
If Seconds2(E) > 10 Then
    picResults.Print LA2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print LA2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1988 olympic games'
Private Sub cmd1988_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Seoul(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Seoul2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1988.txt" For Input As #1
Open App.Path & "\10k1988.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Seoul Results"
For A = 1 To 8
    Input #1, Seoul(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Seoul(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Seoul(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Seoul Results"
For B = 1 To 8
    Input #2, Seoul2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Seoul2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Seoul2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1992 olympic games'
Private Sub cmd1992_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Barcelona(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Barcelona2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1992.txt" For Input As #1
Open App.Path & "\10k1992.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Barcelona Results"
For A = 1 To 8
    Input #1, Barcelona(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Barcelona(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Barcelona(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Barcelona Results"
For B = 1 To 8
    Input #2, Barcelona2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Barcelona2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Barcelona2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 1996 olympic games'
Private Sub cmd1996_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Atlanta(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Atlanta2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k1996.txt" For Input As #1
Open App.Path & "\10k1996.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Atlanta Results"
For A = 1 To 8
    Input #1, Atlanta(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Atlanta(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Atlanta(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Atlanta Results"
For B = 1 To 8
    Input #2, Atlanta2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Atlanta2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Atlanta2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 2000 olympic games'
Private Sub cmd2000_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Sydney(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Sydney2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k2000.txt" For Input As #1
Open App.Path & "\10k2000.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Sydney Results"
For A = 1 To 8
    Input #1, Sydney(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Sydney(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Sydney(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Sydney Results"
For B = 1 To 8
    Input #2, Sydney2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Sydney2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Sydney2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'this button pull up results from 2004 olympic games'
Private Sub cmd2004_Click()
picResults.Cls
'clear the picResults box'
Dim A As Integer, B As Integer
Dim Athens(1 To 8) As Integer, Name(1 To 8) As String, Minute(1 To 8) As Single, Seconds(1 To 8) As Single
Dim D As Integer, E As Integer, Athens2(1 To 8) As Integer, Name2(1 To 8) As String, Minute2(1 To 8) As Single, Seconds2(1 To 8) As Single
Open App.Path & "\5k2004.txt" For Input As #1
Open App.Path & "\10k2004.txt" For Input As #2
'this step opens the 2 data files needed for the arrays'
picResults.Print "5k Athens Results"
For A = 1 To 8
    Input #1, Athens(A), Name(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the arrays needed for display later'
For D = 1 To 8
If Seconds(D) > 10 Then
    picResults.Print Athens(D), Name(D); Tab(30); Minute(D); ":"; FormatNumber(Seconds(D), 2)
Else
    picResults.Print Athens(D), Name(D); Tab(30); Minute(D); ":"; "0"; FormatNumber(Seconds(D), 2)
End If
Next D
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
picResults.Print "10k Athens Results"
For B = 1 To 8
    Input #2, Athens2(B), Name2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the arrays needed for display later'
For E = 1 To 8
If Seconds2(E) > 10 Then
    picResults.Print Athens2(E), Name2(E); Tab(30); Minute2(E); ":"; FormatNumber(Seconds2(E), 2)
Else
    picResults.Print Athens2(E), Name2(E); Tab(30); Minute2(E); ":"; "0"; FormatNumber(Seconds2(E), 2)
End If
Next E
'this step sees if there is a zero in any part of the data, then adds in a 0 if there is one to print it out correctly'
Close 1
Close 2
'these two lines close the data files'
End Sub

'goes back to race results page from olympic results page'
Private Sub cmdBack_Click()
    frmRaceResults.Show
    frmOlympics.Hide
End Sub

'this button should sort out the data from an array and then print it in order from fastest to slowest
Private Sub cmdChampions_Click()
picResults.Cls
Dim TempMinute2(1 To 11) As Single, E As Integer, A As Integer, B As Integer, Champion(1 To 11) As String, Minute(1 To 11) As Integer, Seconds(1 To 11) As Single
Dim Comp As Integer, Pass As Integer, Comp2 As Integer, Pass2 As Integer, Temp As Integer
Dim Temp2 As Single, Temp3 As String, C As Integer, D As Integer, Temp4 As Integer, Temp5 As Single, Temp6 As String
Dim Champion2(1 To 12) As String, Minute2(1 To 12) As Single, Seconds2(1 To 12) As Single
'declare all variables, and all arrays for this button'
Open App.Path & "\best5k.txt" For Input As #1
Open App.Path & "\best10k.txt" For Input As #2
'these two lines open up the 2 arrays needed for my file input'
For A = 1 To 11
    Input #1, Champion(A), Minute(A), Seconds(A)
Next A
'this step inputs the data into the array'
picResults.Print "5k Olympic Champions in order of fastest time to slowest"
For Pass = 1 To 10
    For Comp = 1 To 11 - Pass
        If Seconds(Comp) > Seconds(Comp + 1) Then
            Temp = Minute(Comp)
            Minute(Comp) = Minute(Comp + 1)
            Minute(Comp + 1) = Temp
            Temp2 = Seconds(Comp)
            Seconds(Comp) = Seconds(Comp + 1)
            Seconds(Comp + 1) = Temp2
            Temp3 = Champion(Comp)
            Champion(Comp) = Champion(Comp + 1)
            Champion(Comp + 1) = Temp3
        End If
    Next Comp
Next Pass
'this chunk of steps sorts out the data from slowest time to fastest time'
For C = 1 To 11
If Seconds(C) > 10 Then
    picResults.Print Champion(C), Tab(30); Minute(C); ":"; FormatNumber(Seconds(C), 2)
Else
    picResults.Print Champion(C), Tab(30); Minute(C); ":"; "0"; FormatNumber(Seconds(C), 2)
End If
Next C
'this if/then statement sees if there are zeros in the data, and if there are then it prints them out with a zero instead of nothing and prints out the rest of the data in order from fastest to slowest'
picResults.Print
picResults.Print "10k Olympic Champions in order of fastest time to slowest"
For B = 1 To 12
    Input #2, Champion2(B), Minute2(B), Seconds2(B)
Next B
'this step inputs the data into the array'
For D = 1 To 12
If Seconds2(D) > 10 Then
    picResults.Print Champion2(D), Tab(30); Minute2(D); ":"; FormatNumber(Seconds2(D), 2)
Else
    picResults.Print Champion2(D), Tab(30); Minute2(D); ":"; "0"; FormatNumber(Seconds2(D), 2)
End If
Next D
'this if/then statement sees if there are zeros in the data, and if there are then it prints them out with a zero instead of nothing and prints out the rest of the data in order from fastest to slowest'
Close 1
Close 2
'these 2 steps close the 2 data files
End Sub
