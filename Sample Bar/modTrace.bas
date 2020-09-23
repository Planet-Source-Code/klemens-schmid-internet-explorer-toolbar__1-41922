Attribute VB_Name = "modTrace"
'*********************************************************************************************
' Author:      Klemens Schmid
' E-Mail:      klemens.schmid@gmx.de
' Web Page:    http://www.klemid.de
'
' Description
' This is a preliminary version of the a trace module which will later be replaced by
' a sophisticated trace mechnism.
'*********************************************************************************************

Option Explicit

'defines the trace level
Public Trace As Integer

Public Sub Log(ByVal What$, _
               Optional ByVal ModuleName$ = "", _
               Optional ByVal ProcName$ = "", _
               Optional ByVal CompName$ = "")
'set trace level
Trace = 9
'write a trace log entry
Debug.Print ModuleName, ProcName, What
End Sub

