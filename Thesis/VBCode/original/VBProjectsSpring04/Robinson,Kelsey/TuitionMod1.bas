Attribute VB_Name = "Module1"
' Deciding on a College
' Module1 (TuitionMod1)
' Kelsey Robinson
' March 10th, 2004
' I used this module to dimension all of my variables in one place,
' so that I did not have to dim each variable on each form.

Option Explicit
Public College(1 To 20) As String
Public Tuition(1 To 20) As Single
Public Distance(1 To 20) As Single
Public CTR As Single
Public Pass As Integer
Public Comp As Integer
Public TempCollege As String
Public J As Integer
Public TempTuition As Single
Public TempDistance As Single
Public Found As Boolean
Public position As Integer
Public Willing As Single
Public WillingDist As Single
Public PATH As String
