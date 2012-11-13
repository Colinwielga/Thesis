Attribute VB_Name = "ModuleGeneral"
Option Explicit

'   Day at the Capitol and MN Private College Information Tool
'   Form: ModuleGeneral
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of this module is to declare variables that can be used throughout entire application.
 

Public SchoolName(1 To 20) As String        'Variables used for sorting in AboutMPCC form.
Public enrollment(1 To 20) As Integer
Public tuition(1 To 20) As Integer
Public location(1 To 20) As String
Public tempSchoolName As String, tempenrollment As Integer, temptuition As Double, templocation As String

Public School(1 To 20) As String            'Variables used for match and stop search and sorting in DAC form.
Public Day(1 To 20) As String
Public Registered(1 To 20) As Double
Public TempSchool As String, tempday As String, tempregistered As Double

Public pass As Integer, pos As Integer, j As Integer, ctr As Integer


Public ie As Object                         'Used throughout application for hyperlink to Internet Explorer pages.



