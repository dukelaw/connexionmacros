'MacroName: Add RDA Encoding
'MacroDescription: Adds 040 $b, $e LDR/18 (Descriptive Cataloging Form) to a Bibliographic Record.

' Written by Sean Chen <schen@law.duke.edu>
' Revised: November 13, 2013

Option Explicit
Dim sDesc, s040 As String
Dim sTemp As String
Dim i As Integer
Dim CS As Object

Sub FixStr (sStr$, sFnd$, sFix$, nCase%, nLoop%)
    ' Update a string by inserting a string into another
    If nCase = 0 Then
        If nLoop = 0 Then
            If InStr(UCase(sStr), UCase(sFnd)) <> 0 Then
                sStr = Left(sStr, InStr(UCase(sStr), UCase(sFnd)) - 1) & sFix & Mid(sStr, InStr(UCase(sStr), UCase(sFnd)) + Len(sFnd))
            End If
        Else
            Do While InStr(UCase(sStr), UCase(sFnd)) <> 0
                sStr = Left(sStr, InStr(UCase(sStr), UCase(sFnd)) - 1) & sFix & Mid(sStr, InStr(UCase(sStr), UCase(sFnd)) + Len(sFnd))
            Loop
        End If
    Else
        If nLoop = 0 Then
            If InStr(sStr, sFnd) <> 0 Then
                sStr = Left(sStr, InStr(sStr, sFnd) - 1) & sFix & Mid(sStr, InStr(sStr, sFnd) + Len(sFnd))
            End If
        Else
            Do While InStr(sStr, sFnd) <> 0
                sStr = Left(sStr, InStr(sStr, sFnd) - 1) & sFix & Mid(sStr, InStr(sStr, sFnd) + Len(sFnd))
            Loop
        End If
    End If
End Sub


Sub Main
   Set CS = CreateObject("Connex.Client")
   ' Must be set to top or bottom fixed field display.
   If CS.FixedFieldPosition = 0 Then
       MsgBox "This macro needs the fixed fields to be displayed either on top or on the bottom of the display.", 0, "Convert to RDA Macro"
   End If
   ' Fixed Fields
   CS.SetFixedField "Desc", "i"

   ' Insert language and cataloging code
   CS.GetField "040", 1, s040
   If InStr(s040, "ßc") = False Then
       MsgBox "There is no subfield c in the 040"
   ElseIf InStr(s040, "ßb eng ße rda") >= 0 Then
       MsgBox "This record already has a $b $e"
   Else
       If InStr(s040, "ße rda") >= 0 Then
           FixStr s040, "ße", "ßb eng ße", 0, 0
       ElseIf InStr(s040, "ßb eng") >= 0 Then
           FixStr s040, "ßc", "ßb rda ßc", 0, 0
       Else
           FixStr s040, "ßc", "ßb eng ße rda ßc", 0, 0
       End If
   End If
   CS.SetField 1, s040

End Sub