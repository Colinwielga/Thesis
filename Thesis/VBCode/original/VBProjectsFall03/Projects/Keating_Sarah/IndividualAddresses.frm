VERSION 5.00
Begin VB.Form frmIndividualAddresses 
   BackColor       =   &H00800080&
   Caption         =   "Address"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form2"
   ScaleHeight     =   5355
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   12000
      TabIndex        =   3
      Top             =   10920
      Width           =   3135
   End
   Begin VB.CommandButton cmdEmailForm 
      Caption         =   "Click to View a List of all Email Addresses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   12000
      TabIndex        =   2
      Top             =   7320
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5760
      ScaleHeight     =   3555
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   2880
      Width           =   9015
   End
   Begin VB.CommandButton cmdIndividual 
      Caption         =   "Click to Enter a Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   1080
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
   End
End
Attribute VB_Name = "frmIndividualAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmIndividualAddresses(IndividualAddresses.frm)
'Sarah Keating
'10-22-03
'Purpose: This form allows the user to input a name and view that person's address,
'           phone number, and email address.  It also provides a link to the next form.
Option Explicit
Dim Names(1 To 22) As String, Address(1 To 22) As String, PhoneNumber(1 To 22) As String, Email(1 To 22) As String
' Again, I had trouble getting the form load function to process correctly,
' so I dimensioned all arrays at this point.
Dim FirstLastNames(1 To 22) As String


Private Sub cmdEmailForm_Click()
frmIndividualAddresses.Hide
frmEmail.Show
' Allows the user to go to the next form

End Sub

Private Sub cmdIndividual_Click()
picResults.Cls
' Clears the picturebox


Open PATH & "Addresses.txt" For Input As #1

Dim I As Integer
Dim N As String
Dim NotFound As Boolean
Do Until EOF(1)
    I = I + 1
    Input #1, Names(I), Address(I), PhoneNumber(I), Email(I), FirstLastNames(I)
Loop
' Inputs all of the information into arrays
Close

N = InputBox("Enter a name to look up an address and display a picture")
' The InputBox allows the user to enter an individual name and look up that person's information.

I = 0
NotFound = True
Do While NotFound And I < 22
' The sequential search goes through the list until it finds the name that was entered by the user

    I = I + 1
    If N = FirstLastNames(I) Then NotFound = False
Loop
If NotFound Then
        picResults.Print "That person is not in the address book"
    Else
        picResults.Print FirstLastNames(I)
        picResults.Print Address(I)
        picResults.Print PhoneNumber(I)
        picResults.Print Email(I)
        ' Prints the desired output
        
End If

    
    

End Sub


Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
End Sub
