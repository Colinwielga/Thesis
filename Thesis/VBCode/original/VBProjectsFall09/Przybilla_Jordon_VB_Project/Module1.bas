Attribute VB_Name = "Module1"
Option Explicit
'this module will contain all arrays used in the program so no extra load buttons are needed

Public Terms(1 To 50) As String, Ctr As Integer 'this will allow the file to be read for the next form without extra buttons'
'frmTerms array

Public Facts(1 To 10) As String 'frmVehicles array, facts about deer-vehicle collisions

Public AvoidTips(1 To 10) As String 'frmVehicles array, tips for avoiding collisions

Public x As Integer 'used to print in several forms

Public CantAvoid(1 To 10) As String 'frmVehicles array, tips for after you hit a deer

Public collisionslides(1 To 6) As String

Public rifleregs(1 To 10) As String

Public rifletips(1 To 10) As String

Public info(1 To 10) As String

Public Guns(1 To 10) As String, Caliber(1 To 10) As Single, Grain(1 To 10) As Long, Energy(1 To 10) As Long
