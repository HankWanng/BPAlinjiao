VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=====================================================================
' © Copyright Yokogawa Electric Corporation 2010
' All rights reserved.  This software is supplied for a single
' computer system and is not to be copied, amended, or translated
' into another program language without the prior written consent
' of Yokogawa Electric Corporation.
'
'=====================================================================

Option Explicit
Option Base 1
Option Compare Text
DefLng A-Z

'
' If double clicking on an editable cell displays the Edit dialogue
'
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
On Error GoTo Error
    If Target.Count = 1 And Application.Run("'Exaquantum Report Add-In.xla'!CellWriteable") Then
        Application.Run ("'Exaquantum Report Add-In.xla'!EditTagValue")
    End If
    
    If Target.MergeCells And Application.Run("'Exaquantum Report Add-In.xla'!CellWriteable") Then
        Application.Run ("'Exaquantum Report Add-In.xla'!EditTagValue")
    End If
    
EndRoutine:
    Cancel = True
    Exit Sub
Error:
    Resume EndRoutine
End Sub

