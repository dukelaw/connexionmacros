' MacroName: SearchCallNumber
' MacroDescription: Search for a Call Number in Classification Web and inside configured shelflists (Library of Congress and Duke University)
' Caveat: You may need to be auto-logged into Classification Web for the search to work.

Dim sStr$

Sub FixStr (sStr$, sFnd$, sFix$, nCase%, nLoop%)
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
   Dim CS As Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0
   If CS  Is Nothing Then
   Set CS  = CreateObject("Connex.Client")
   End If
   Dim nBool%, nRow%
   Dim sHdg$, sLab$, sTag$
   Dim sCallNumber$, sClassNumber$
   Dim sLC$, sDuke$, sClassWeb$
   Dim sBrowser$
   sBrowser = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
   ' example: https://search.catalog.loc.gov/browse/call-numbers?itemsPerPage=25&search=KH2020&shelfKey=KH2020
   ' search LC shelflist
   sLC = " https://search.catalog.loc.gov/browse/call-numbers?itemsPerPage=25&search="
   ' Duke example: https://find.library.duke.edu/?search_field=call_number&q=HGV6022.B6
   sDuke= " ""https://find.library.duke.edu/?search_field=call_number&q="

   sClassWeb = " ""https://classweb.org/min/minaret?app=Corr&mod=Search&menu=/Menu/&iname=l2nh&iterm="
   sClassWebSuffix = "&ilabel=LC%20class%20%23%20--%3E%20LCSH%20(w%2Fnames)"

   nRow = CS.CursorRow
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or _
      CS.ItemType = 3 Or CS.ItemType = 4 or CS.ItemType = 14 or _
      CS.ItemType = 17 or CS.ItemType = 18 or CS.ItemType = 19 or _
      CS.ItemType = 20 or CS.ItemType = 26 or CS.ItemType = 31 or _
      CS.ItemType = 33 or CS.ItemType = 35 Then
      If CS.GetFieldLine(nRow, sCallNumber) = True Then
         sTag = Mid$(sCallNumber, 1, 3)
         sCallNumber = Mid$(sCallNumber, 6)
         ' Get just class number
         sClassNumber = Mid$(sCallNumber, 1, InStr(sCallNumber, "ß") - 1)
         FixStr sClassNumber, " ",  Chr(37) + "20", 0, 1
         Do While InStr(sCallNumber, "ß") <> 0
               sCallNumber = Mid$(sCallNumber, 1, InStr(sCallNumber, "ß") - 1) & Mid$(sCallNumber, InStr(sCallNumber, "ß") + 3)
         Loop
         FixStr sCallNumber, " ",  Chr(37) + "20", 0, 1
         sLC = sLC & sCallNumber & """"
         sDuke = sDuke & sClassNumber & """"
         sClassWeb = sClassWeb & sCallNumber & sClassWebSuffix & """"
         Shell(sBrowser & sLC)
         Shell(sBrowser & sClassWeb)
         Shell(sBrowser & sDuke)
      Else
         MsgBox "Place the cursor in an 050 or 090 field to shelflist", 0, "Search Call Number Macro"
      End If

    Else
        Msgbox "An authority or bibliographic record must be displayed in order to use this macro", 0, "Search Call Number Macro"
    End If
End Sub
