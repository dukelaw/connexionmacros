'MacroName: ClassSubjectCorrelations
'MacroDescription: A macro to lookup correlations between Library of Congress Classification and Library of Congress Subject Headings

Sub Main

   Dim Client As Object
   On Error Resume Next
   Set Client  = GetObject(,"Connex.Client")
   On Error GoTo 0
   If Client  Is Nothing Then
      Set Client  = CreateObject("Connex.Client")
   End If
   Dim sFieldData As String
   Dim sClassweb As String
   Dim nRow As Integer
   Dim sLab As String
   Dim sTag as String
   Dim sSpace As String
   Dim x
   Dim i
   sBrowser = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""

   sClassWebLCC = sBrowser + "https://classweb.org/min/minaret?app=Corr&mod=Search&menu=/Menu/&iname=l2nh&iterm="
   sClassWebLCCSuffix = "&ilabel=LC%20class%20%23%20--%3E%20LCSH%20(w%2Fnames)"
   sClassWebLCSH = sBrowser + "https://classweb.org/min/minaret?app=Corr&mod=Search&menu=/Menu/&iname=sh2l&iterm="
   sClassWebLCSHSuffix = "&ilabel=LCSH%20--%3E%20LC%20class%20%23"

   On Error Resume Next
   Set Client  = GetObject(,"Connex.Client")
   On Error GoTo 0
   If Client  Is Nothing Then
      Set Client  = CreateObject("Connex.Client")
   End If
      nRow = Client.CursorRow
   If Client.IsOnline = False Then
      Client.Logon "","",""
   End If

   If Client.ItemType = 0 or Client.ItemType = 1 or Client.ItemType = 2 or Client.ItemType = 3 or Client.ItemType = 4 or Client.ItemType = 14 or Client.ItemType = 17 or Client.ItemType = 18 or Client.ItemType = 19 or Client.ItemType = 20 Then
   If Client.GetFieldLine (nRow, sFieldData) = True Then
      sTag = Mid$(sFieldData, 1, 3)
      sFieldData = Mid$(sFieldData, 6)
      If InStr ("050, 055, 090, 099", sTag) <> 0 Then
         If InStr (sFieldData, "ßa") <> 0 Then
           sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßa") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßa") + 3)
           End If
           if InStr (sFieldData, "ßb") <> 0 Then
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßb") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßb") + 3)
           End If
           if InStr (sFieldData, "ßd") <> 0 Then
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßd") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßd") + 3)
           End If
           if InStr (sFieldData, "ß3") <> 0 Then
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ß3") - 1) & Mid$(sFieldData, InStr(sFieldData, "ß3") + 3)
           End If
           Do While InStr(sFieldData, " ") <> 0
            sFieldData= Mid$(sFieldData, 1, InStr(sFieldData, " ") -1) & "%20" & Mid$(sFieldData, InStr(sFieldData, " ") + 1)
           Loop
           sClassWebLCC = sClassWebLCC & sFieldData & sClassWebLCCSuffix
           x = shell (sClassWebLCC, 1)
      ElseIf InStr("650", sTag) <> 0 Then
           Do While InStr(sFieldData, "ßx") <> 0
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßx") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßx") + 3)
           Loop
           Do While InStr(sFieldData, "ßy") <> 0
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßy") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßy") + 3)
           Loop
           Do While InStr(sFieldData, "ßz") <> 0
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßz") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßz") + 3)
           Loop
           Do While InStr(sFieldData, "ßv") <> 0
            sFieldData = Mid$(sFieldData, 1, InStr(sFieldData, "ßv") - 1) & Mid$(sFieldData, InStr(sFieldData, "ßv") + 3)
           Loop
           Do While InStr(sFieldData, " ") <> 0
            sFieldData= Mid$(sFieldData, 1, InStr(sFieldData, " ") -1) & "+" & Mid$(sFieldData, InStr(sFieldData, " ") + 1)
           Loop
           sClassWebLCSH = sClassWebLCSH & sFieldData & sClassWebLCSHSuffix
           x =shell (sClassWebLCSH, 1)
      Else
           MsgBox("Please select a LC Call Number field or a Topical Subject heading field.")
      End If

  End If
  End If

End Sub
