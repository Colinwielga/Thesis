VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H000000C0&
   Caption         =   "More Information about Michelangelo Art"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   Picture         =   "frmInfo.frx":0000
   ScaleHeight     =   7695
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Click here to Exit the Program"
      Height          =   1155
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
   End
   Begin VB.PictureBox pbxResults 
      Height          =   6015
      Left            =   3720
      ScaleHeight     =   5955
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   480
      Width           =   5175
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Click here to show more information about a specific work."
      Height          =   1155
      Left            =   480
      Picture         =   "frmInfo.frx":2C0D
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Click here to sort the works by the date they were started."
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Click here to show Standard information about each picture. "
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Michelangelo Art
'Info(frmInfo.frm)
'Beth Welle
'October 29, 2003
'Purpose of this form is to show the user more information about each work, as well as to sort the
'informationto see which works will never be completed.




'makes a table of the information by going through the array.
Private Sub cmdPrint_Click()
pbxResults.Cls
pbxResults.Print "Michelangelo's Works:"
pbxResults.Print
pbxResults.Print "Name", "Year", "Completed Work"
pbxResults.Print "-------------------------------------------------------------------"
For i = 1 To 6
    pbxResults.Print Work(i), Year(i), Complete(i)
Next i

End Sub

'ends the program so the user may exit
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdShow_Click()
Dim w As String

'asks the user for name of the artwork, and then prints out the related information

pbxResults.Cls
w = InputBox("Please enter the name of a work you would like to find more information about.")

Select Case w
    Case "Atlas"
        pbxResults.Print "This work is an unfinished statue that was intended"
        pbxResults.Print "most likely to be a part of Pope Julius II's tomb."
        pbxResults.Print "One of four, this work is in the Accademia in Florence."
    Case "Crucifix"
        pbxResults.Print "This work is a wooden crucifix and is the only known"
        pbxResults.Print "work by Michelangelo in wood. The figure of Christ is"
        pbxResults.Print "nude (which scandalized his contemporaries); throughout"
        pbxResults.Print "his career Michelangelo created his masterpieces as a"
        pbxResults.Print "celebration of the human body."
    Case "Moses"
        pbxResults.Print "This statue is the completed figure that was designed"
        pbxResults.Print "for Pope Julius II's tomb. It shows an imposing, seated"
        pbxResults.Print "Moses and is located in San Pietro in Vincoli in Rome."
    Case "Pieta"
        pbxResults.Print "One of Michelangelo's most famous works, the Pietà is"
        pbxResults.Print "located in St. Peter's basilica in Rome. The face of the"
        pbxResults.Print "Virgin is especially youthful; that of Christ shows no"
        pbxResults.Print "sign of pain."
    Case "Rondanini"
        pbxResults.Print "This statue is a work that Michelangelo was probably"
        pbxResults.Print "working on shortly before his death. Originally a larger"
        pbxResults.Print "work, Michelangelo cut away portions of the original"
        pbxResults.Print "statue and began carving again; the result is an unfinished"
        pbxResults.Print "statue of the standing figures of Christ and his mother."
    Case "Madonna"
        pbxResults.Print "Known as the Bruges Madonna, this statue was carved by"
        pbxResults.Print "Michelangelo at about the same time that he was working"
        pbxResults.Print "on the colossal David. This work shows a seated Madonna"
        pbxResults.Print "with a nude Christ child standing between her knees."
    Case Else
        MsgBox "Please check the spelling of the artwork you wish to learn more about", , "Error"
End Select

End Sub

Private Sub cmdSort_Click()
pbxResults.Cls
pbxResults.Print "Michelangelo's Works"
pbxResults.Print
pbxResults.Print "Name", "Year", "Completed Work"
pbxResults.Print "-------------------------------------------------------------------"

'sorts through the array, if the years compared are not in the correct order,(smallest to biggest)
'then it will switch the years and their related information

Dim N As Integer
N = 6
Dim pass As Integer
Dim temp As String

For pass = 1 To N - 1
    For i = 1 To N - pass
        If Year(i) > Year(i + 1) Then
            temp = Year(i)
            Year(i) = Year(i + 1)
            Year(i + 1) = temp
            
            temp = Work(i)
            Work(i) = Work(i + 1)
            Work(i + 1) = temp
            
            temp = Complete(i)
            Complete(i) = Complete(i + 1)
            Complete(i + 1) = temp
        End If
    Next i
Next pass

For i = 1 To 6
    pbxResults.Print Work(i), Year(i), Complete(i)
Next i

End Sub

'loads the file where the data is stored, and puts it into an array

Private Sub Form_Load()

strPath = "N:\CS130\handin\BethWelle\vbProject\"
strFile = strPath & "Art.txt"
Open strFile For Input As #1
For i = 1 To 6
    Input #1, Work(i), Year(i), Complete(i)
Next i

End Sub
