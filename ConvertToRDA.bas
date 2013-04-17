' MacroName: Convert to RDA
' MacroDescription: Converts an existing record into an PCC RDA record with
' additional RDA Fields.

' Written by Sean Chen <schen@law.duke.edu>
' Revised: April 16, 2013


Option Explicit
Dim sELvl, s040, s042, s260, s264_1, s264_4, s300, s504, s336, s337, s338 As String
Dim sType, sBLvl, sForm, sFMus, sPubDate, sCopyRightDate, sCont, sIlls As String
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
   'CS.GetFixedField ("ELvl", sELvl)
   ' MsgBox "Encoding Level = " + sELvl, 0, "Testing"
   CS.SetFixedField "ELvl", " "
   CS.SetFixedField "Srce", "c"
   CS.SetFixedField "Desc", "i"

   ' Insert language and cataloging code
   CS.GetField "040", 1, s040
   If InStr(s040, "ßb eng ße rda") = False Then
       If InStr(s040, "ßb") Then
           ' Fix 
           s040 = Left(s040, InStr(s040, "ßb") - 1) & Mid(s040, InStr(s040, "ßb") + 7)
       End If
       FixStr s040, "ßc", "ßb eng ße rda ßc", 0, 0
   End If
   CS.SetField 1, s040

   ' Add a pcc authentication code
   If CS.GetField("042", 1, s042) = True And InStr(s042, "pcc") = False Then
       s042 = s042 + " ßa pcc"
       CS.SetField 1, s042
   End If
   If InStr(s042, "pcc") = False Then
       s042 = "042  pcc"
       CS.AddField 1, s042
   End If

   ' Transform 260
   If CS.GetField("260", 1, s260) = True Then
       s264_1 = s260
       FixStr s264_1, "260  ", "264 1", 0, 0
       If InStr(s260, "ßc c") Then
           'Add a copyright field
           sPubDate = Mid(s264_1, InStr(s264_1, "ßc") + 4, 4)
           sCopyRightDate = Mid(s264_1, InStr(s264_1, "ßc") + 4, 4)
           s264_1 = Left(s264_1, InStr(s264_1, "ßc") + 2) & "[" & sPubDate & "]"
           s264_4 = "264 4ßc " & Chr(202) & sCopyRightDate
           CS.AddField 1, s264_1
           CS.AddField 2, s264_4
           'MsgBox "Copyright: " + s264_4, 0, "Testing"
           CS.DeleteField "260", 1
           CS.SetFixedField "DtSt", "t"
           CS.SetFixedField "Dates", sPubDate
           ' use "," for Date2
           CS.SetFixedField ",", sCopyRightDate
       Else
           CS.DeleteField "260", 1
           CS.AddField 1, s264_1
       End If
   End If
   ' Transform 300 First one only?
   ' Can refactor later if we need it a third time
   If CS.GetField("300", 1, s300) = True And InStr(s300, " p. ") > 0 Then
       FixStr s300, " p. ", " pages ", 0 , 0
       ' MsgBox "Pages?: " + s300, 0, "Testing"
       CS.SetField 1, s300
   End If
   ' Illustrations
   If CS.GetField("300", 1, s300) = True And InStr(s300, "ill.") > 0 Then
       FixStr s300, "ill.", "illustrations", 0 , 0
       CS.SetField 1, s300
   End If

   ' Fixed Field for Illustrations
   CS.GetFixedField "Ills", sIlls
   If InStr(s300, "illustrations") And InStr(sIlls, "a") = 0 Then
       CS.SetFixedField "Ills", "a" & Mid(sIlls, 1, 3)
   End If
   
   ' Transform 504
   CS.GetFixedField "Cont", sCont
   If CS.GetField("504", 1, s504) = True And InStr(s504, "(p. ") > 0 Then
       MsgBox s504
       FixStr s504, "(p. ", "(pages ", 0 , 0
       if InStr(s504, "[") > 0 Or InStr(s504, "]") Then
          FixStr s504, "[", "", 0, 0
          FixStr s504, "]", "", 0, 0
       End If
       CS.SetField 1, s504
   End If

   ' Fixed field for Bibliography
   If CS.GetField("504", 1, s504) = True and InStr(sCont, "b") = 0 Then
       CS.SetFixedField "Cont", "b" & Mid(sCont, 1, 3)
   End If
   ' Index for Bibliography
   If CS.GetField("504", 1, s504) = True and InStr(s504, "index") > 0 Then
       CS.SetFixedField "Indx", "1"
   End If

   ' Fixed field for Indexes

   
   'Insert 336, 337, 338'
   ' Can do this heuristically?
   CS.GetFixedField "Type", sType 
   CS.GetFixedField "BLvl", sBLvl
   CS.GetFixedField "Form", sForm
   If sType = "a" And (sForm = "" Or (sForm = "r" Or sForm="d")) Then
       'tangible book
       s336 = "336  text ß2 rdacontent"
       s337 = "337  unmediated ß2 rdamedia"
       s338 = "338  volume ß2 rdacarrier"
   ElseIf sType = "a" And (sForm = "o") Then
       'online monograph or serial
       s336 = "336  text ß2 rdacontent"
       s337 = "337  computer ß2 rdamedia"
       s338 = "338  online resource ß2 rdacarrier"
   Else       
       s336 = "336  {UNABLE_TO_EXTRACT} ß2 rdacontent"
       s337 = "337  {UNABLE_TO_EXTRACT} ß2 rdamedia"
       s338 = "338  {UNABLE_TO_EXTRACT} ß2 rdacarrier"
   End If
   CS.AddField 1 , s336
   CS.AddField 1 , s337
   CS.AddField 1 , s338
   
End Sub