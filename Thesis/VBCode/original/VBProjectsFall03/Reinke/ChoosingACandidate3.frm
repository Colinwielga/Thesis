VERSION 5.00
Begin VB.Form frmyourcandidate 
   BackColor       =   &H000000FF&
   Caption         =   "Your Candidate"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxflag 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   5520
      ScaleHeight     =   2235
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   1320
      Width           =   4215
   End
   Begin VB.PictureBox pbxcand2name 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   8160
      Width           =   2055
   End
   Begin VB.PictureBox pbxcand1name 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.PictureBox pbxcand2 
      BackColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox pbxcand1 
      BackColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdcand 
      BackColor       =   &H00FF0000&
      Caption         =   "Show Candidate"
      Height          =   2055
      Left            =   360
      TabIndex        =   3
      Top             =   9000
      Width           =   4095
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   9240
      TabIndex        =   2
      Top             =   9840
      Width           =   2415
   End
   Begin VB.CommandButton cmdstartover 
      Caption         =   "Start over"
      Height          =   1095
      Left            =   6480
      TabIndex        =   1
      Top             =   9840
      Width           =   2415
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00FF8080&
      Height          =   2175
      Left            =   4440
      ScaleHeight     =   2115
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   5160
      Width           =   6495
   End
End
Attribute VB_Name = "frmyourcandidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChoosingACandidate(ChoosingACandidate)
'frm your candidate(frmyourcandidate)
'Elaina Reinke
'October 30,2003
'this form adds up the sum of all the options in all the other forms
'then it prints them out to a picturebox and will print out a picture
'of the candidate or candidates that share values most similar to
'your values


Private Sub cmdcand_Click()
'clears the pictureboxes everytime the program is run
pbxresults.Cls
pbxcand1.Cls
pbxcand2.Cls
pbxcand1name.Cls
pbxcand2name.Cls
'this adds up the sum of all the values together and figures out which
'candidate has values most similar to yours according to the way you
'answered the questions and also prints the picture of the candidate
'path is used as a variable so that we do not need to type in the entire command everytime

If sum = 10 Then
    pbxresults.Print "The candidate with values most similar to yours is Ralph Nader"
    pbxcand1.Picture = LoadPicture(PATH & "nader.jpg")
    pbxcand1name.Print "Ralph Nader"
    ElseIf sum > 10 And sum <= 19 Then
            pbxresults.Print "The candidates with values most similar to yours are Ralph Nader and Dennis Kucinich"
            pbxcand1.Picture = LoadPicture(PATH & "nader.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Dennis_Kucinich.jpg")
            pbxcand1name.Print "Ralph Nader"
            pbxcand2name.Print "Dennis Kucinich"
        ElseIf sum = 20 Then
            pbxresults.Print "The candidate with values most similar to yours is Dennis Kucinich"
            pbxcand1.Picture = LoadPicture(PATH & "Dennis_Kucinich.jpg")
            pbxcand1name.Print "Dennis Kucinich"
        ElseIf sum > 20 And sum <= 29 Then
            pbxresults.Print "The candidates with values most similar to yours are Dennis Kucinich and Rev. Al Sharpton"
            pbxcand1.Picture = LoadPicture(PATH & "Dennis_Kucinich.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Al_Sharpton.jpg")
            pbxcand1name.Print "Dennis Kucinich"
            pbxcand2name.Print "Rev. Al Sharpton"
        ElseIf sum = 30 Then
            pbxresults.Print "The candidate with values most similar to yours is Rev. Al Sharpton"
            pbxcand1.Picture = LoadPicture(PATH & "Al_Sharpton.jpg")
            pbxcand1name.Print "Rev. Al Sharpton"
        ElseIf sum > 30 And sum <= 39 Then
            pbxresults.Print "The candidates with values most similar to yours are Rev. Al Sharpton and John Kerry"
            pbxcand1.Picture = LoadPicture(PATH & "Al_Sharpton.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "John_Kerry.jpg")
            pbxcand1name.Print "Rev. Al Sharpton"
            pbxcand2name.Print "John Kerry"
        ElseIf sum = 40 Then
            pbxresults.Print "The candidate with values most similar to yours is John Kerry"
            pbxcand1.Picture = LoadPicture(PATH & "John_Kerry.jpg")
            pbxcand1name.Print "John Kerry"
        ElseIf sum > 40 And sum <= 49 Then
            pbxresults.Print "The candidates with values most similar to yours are John Kerry and Howard Dean"
            pbxcand1.Picture = LoadPicture(PATH & "John_Kerry.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Howard_Dean.jpg")
            pbxcand1name.Print "John Kerry"
            pbxcand2name.Print "Howard Dean"
        ElseIf sum = 50 Then
            pbxresults.Print "The candidate with values most similar to yours is Howard Dean"
            pbxcand1.Picture = LoadPicture(PATH & "Howard_Dean.jpg")
            pbxcand1name.Print "Howard Dean"
        ElseIf sum > 50 And sum <= 59 Then
            pbxresults.Print "The candidates with values most similar to yours are Howard Dean and Wesley Clark"
            pbxcand1.Picture = LoadPicture(PATH & "Howard_Dean.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Wesley_Clark.jpg")
            pbxcand1name.Print "Howard Dean"
            pbxcand2name.Print "Wesley Clark"
        ElseIf sum = 60 Then
            pbxresults.Print "The candidate with values most similar to yours is Wesley Clark"
            pbxcand1.Picture = LoadPicture(PATH & "Wesley_Clark.jpg")
            pbxcand1name.Print "Wesley Clark"
        ElseIf sum > 60 And sum <= 69 Then
            pbxresults.Print "The candidates with values most similar to yours are Wesley Clark and John Edwards"
            pbxcand1.Picture = LoadPicture(PATH & "Wesley_Clark.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "John_Edwards.jpg")
            pbxcand1name.Print "Wesley Clark"
            pbxcand2name.Print "John Edwards"
        ElseIf sum = 70 Then
            pbxresults.Print "The candidate with values most similar to yours is John Edwards"
            pbxcand1.Picture = LoadPicture(PATH & "John_Edwards.jpg")
            pbxcand1name.Print "John Edwards"
        ElseIf sum > 70 And sum <= 79 Then
            pbxresults.Print "The candidates with values most similar to yours are John Edwards and Carol Moseley-Braun"
            pbxcand1.Picture = LoadPicture(PATH & "John_Edwards.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Carol_Moseley-Braun.jpg")
            pbxcand1name.Print "John Edwards"
            pbxcand2name.Print "Carol Moseley-Braun"
        ElseIf sum = 80 Then
            pbxresults.Print "The candidate with values most similar to yours is Carol Moseley-Braun"
            pbxcand1.Picture = LoadPicture(PATH & "Carol_Moseley-Braun.jpg")
            pbxcand1name.Print "Carol Moseley-Braun"
        ElseIf sum > 80 And sum <= 89 Then
            pbxresults.Print "The candidates with values most similar to yours are Carol Moseley-Braun and Joe Lieberman"
            pbxcand1.Picture = LoadPicture(PATH & "Carol_Moseley-Braun.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Joseph_Lieberman.jpg")
            pbxcand1name.Print "Carol Moseley-Braun"
            pbxcand2name.Print "Joe Lieberman"
        ElseIf sum = 90 Then
            pbxresults.Print "The candidate with values most similar to yours is Joe Lieberman"
            pbxcand1.Picture = LoadPicture(PATH & "Joseph_Lieberman.jpg")
            pbxcand1name.Print "Joe Lieberman"
        ElseIf sum > 90 And sum <= 99 Then
            pbxresults.Print "The candidates with values most similar to yours are Joe Lieberman and Bob Graham"
            pbxcand1.Picture = LoadPicture(PATH & "Joseph_Lieberman.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Bob_Graham.jpg")
            pbxcand1name.Print "Joe Lieberman"
            pbxcand2name.Print "Bob Graham"
        ElseIf sum = 100 Then
            pbxresults.Print "The candidate with values most similar to yours is Bob Graham"
            pbxcand1.Picture = LoadPicture(PATH & "Bob_Graham.jpg")
            pbxcand1name.Print "Bob Graham"
        ElseIf sum > 100 And sum <= 109 Then
            pbxresults.Print "The candidates with values most similar to yours are Bob Graham and Dick Gephardt"
            pbxcand1.Picture = LoadPicture(PATH & "Bob_Graham.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "Dick_Gephardt.jpg")
            pbxcand1name.Print "Bob Graham"
            pbxcand2name.Print "Dick Gephardt"
        ElseIf sum = 110 Then
            pbxresults.Print "The candidate with values most similar to yours is Dick Gephardt"
            pbxcand1.Picture = LoadPicture(PATH & "Dick_Gephardt.jpg")
            pbxcand1name.Print "Dick Gephardt"
        ElseIf sum > 110 And sum <= 119 Then
            pbxresults.Print "The candidates with values most similar to yours are Dick Gephardt and George W. Bush"
            pbxcand1.Picture = LoadPicture(PATH & "Dick_Gephardt.jpg")
            pbxcand2.Picture = LoadPicture(PATH & "George_W__Bush.jpg")
            pbxcand1name.Print "Dick Gephardt"
            pbxcand2name.Print "George W. Bush"
        ElseIf sum = 120 Then
            pbxresults.Print "The candidate with values most similar to yours is George W. Bush"
            pbxcand1.Picture = LoadPicture(PATH & "George_W__Bush.jpg")
            pbxcand1name.Print "George W. Bush"
    End If
End Sub

Private Sub cmdquit_Click()
'quit the program
End
End Sub

Private Sub cmdstartover_Click()
'go back to the beginning of the program
frmyourcandidate.Hide
frmabortion.Show
End Sub


Private Sub Form_Load()
PATH = PATH & "candidatephotos.jpg\"
pbxflag.Picture = LoadPicture(PATH & "americanflag.jpg")


End Sub
