Attribute VB_Name = "Namen"
Option Explicit
'Decalse these variables for the entire project.
'This allows every form to be able to use and read should they need to.
'Since my program runs on case files i have a lot of strings.
'and a total of 4 files.
Public nam As String
Public CaseFile(1 To 500) As String
Public Rapetypes(1 To 100) As String
Public orgdis(1 To 100) As String
Public sitpedo(1 To 100) As String
Public prepedo(1 To 100) As String
Public avoid(1 To 50) As String
Public check(1 To 200) As Integer
Public rape1answer As String
Public rape2answer As String
Public pedo1answer As String
Public pedo2answer As String
