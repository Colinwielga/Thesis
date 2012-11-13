VERSION 5.00
Begin VB.Form frmschedule 
   Caption         =   "Schedule"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   Picture         =   "Schedule Form.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearchschedule 
      Caption         =   "Search UFC 2010 Schedule Fight Number"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7560
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
End
Attribute VB_Name = "frmschedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdgoback_Click()
frmschedule.Hide
frmmainscreen.Show
End Sub

Private Sub cmdsearchschedule_Click()

Dim found As Boolean, ctr As Integer, dateofevent As String, xvenue As String, xlocation As String
Dim datab As Database, records As Recordset2, fightnumber As Integer, eventx As String

Set datab = OpenDatabase(App.Path & "\UFCdatabase.accdb") 'telling program where to draw its data from
Set records = datab.OpenRecordset("UFCschedule") 'telling program where to look in the database
  
    fightnumber = InputBox("Enter the fight number you would like to search.")
    found = False
    ctr = 0
Do While Not records.EOF 'telling program to search database
        ctr = ctr + 1
    If (records![Fight Number]) = (fightnumber) Then
        dateofevent = (records![Date of Event])
        xvenue = (records![venue]) 'defining variables that correspond with those in the database
        xlocation = (records![location])
        eventx = (records![Event])
        found = True
        MsgBox "UFC fight number " & fightnumber & " is " & eventx & " on " & dateofevent & " at " & xvenue & " venue in " & xlocation & "."
    End If
        records.MoveNext
Loop

    If found = False Then
        MsgBox "Fight Not Found"
    End If
End Sub
