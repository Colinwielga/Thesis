VERSION 5.00
Begin VB.Form frmAddresses 
   BackColor       =   &H00808000&
   Caption         =   "Addresses, Phone Numbers, and Email"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   13710
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
      Height          =   2175
      Left            =   14880
      TabIndex        =   4
      Top             =   10920
      Width           =   2895
   End
   Begin VB.CommandButton cmdNextForm 
      Caption         =   "Click to Look Up an Individual Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   14880
      TabIndex        =   3
      Top             =   8280
      Width           =   2895
   End
   Begin VB.PictureBox picResults2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12495
      Left            =   9480
      ScaleHeight     =   12435
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.PictureBox picResults1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12495
      Left            =   3840
      ScaleHeight     =   12435
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdGroupAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to Display a Group List of Family and Friends' Addresses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   4320
      Width           =   2535
   End
End
Attribute VB_Name = "frmAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Address Book
'frmAddresses(Addresses.frm)
'Sarah Keating
'10-21-03
'Purpose: The idea for this project came about because I wanted a way
'           to keep track of family and friends while studying abroad
'           in Ireland.  Rather than using a typical address book, I
'           thought it might be interesting to compile all of the
'           information in a Visual Basic format.  This first form
'           displays a list of all of my contacts with their addresses,
'           phone numbers, and email addresses.  The list is displayed in
'           alphabetical order, and it also provides a link to the next form.



Option Explicit


Dim Names(1 To 22) As String, Address(1 To 22) As String, PhoneNumber(1 To 22) As String, Email(1 To 22) As String
' I had trouble getting the form load function to process correctly,
' so I dimensioned all arrays at this point.
Dim FirstLastNames(1 To 22) As String


Private Sub cmdGroupAddress_Click()
picResults1.Cls
picResults2.Cls
' Clears the picture boxes for multiple uses

Open PATH & "Addresses.txt" For Input As #1
' I was having difficulties declaring this file as a path, so I had
' to open it explicitly each time.


Dim CTR As Integer
Dim Pass As Integer, Comp As Integer, Temp As String


Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Names(CTR), Address(CTR), PhoneNumber(CTR), Email(CTR), FirstLastNames(CTR)
    ' The name, address, phone number, and email address of each person
    ' are put into arrays
Loop
Close

For Pass = 1 To 21
    ' This is the beginning of the alphabetical sort
    
    For CTR = 1 To 22 - Pass
        If Names(CTR) > Names(CTR + 1) Then
            Temp = Names(CTR)
            Names(CTR) = Names(CTR + 1)
            Names(CTR + 1) = Temp
            Temp = Address(CTR)
            Address(CTR) = Address(CTR + 1)
            Address(CTR + 1) = Temp
            Temp = PhoneNumber(CTR)
            PhoneNumber(CTR) = PhoneNumber(CTR + 1)
            PhoneNumber(CTR + 1) = Temp
            Temp = Email(CTR)
            Email(CTR) = Email(CTR + 1)
            Email(CTR + 1) = Temp
            Temp = FirstLastNames(CTR)
            FirstLastNames(CTR) = FirstLastNames(CTR + 1)
            FirstLastNames(CTR + 1) = Temp
        End If
    Next CTR
Next Pass

For CTR = 1 To 11
    picResults1.Print Names(CTR)
    picResults1.Print Address(CTR)
    picResults1.Print PhoneNumber(CTR)
    picResults1.Print Email(CTR)
    picResults1.Print
Next CTR

' The information had to be separated into two picture boxes because it could not
' all be displayed in one.

For CTR = 12 To 22
    picResults2.Print Names(CTR)
    picResults2.Print Address(CTR)
    picResults2.Print PhoneNumber(CTR)
    picResults2.Print Email(CTR)
    picResults2.Print
Next CTR
    ' Prints the information for each person, separated by a blank line
    

End Sub

Private Sub cmdNextForm_Click()
    frmIndividualAddresses.Show
    frmAddresses.Hide
    ' Allows the User to move on to the next form
    
End Sub

Private Sub cmdQuit_Click()
    End
    ' Allows the user to quit the program
    
End Sub
