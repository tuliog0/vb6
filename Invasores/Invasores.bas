Attribute VB_Name = "Module1"
Option Explicit

Declare Function sndPlaySound Lib "MMSystem" (ByVal Ipsound As String, ByVal flag As Integer) As Integer

Global Const SND_ASYNC = &H1
