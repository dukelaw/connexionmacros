'MacroName: Add Database Notes
'MacroDescription: Adds database notes to a record
'
' Written by Sean Chen <schen@law.duke.edu>
' Revised: October 23, 2014
'****************************************************************************************

Option Explicit

Function DateString() As String
    Dim sDateStr$, sDay$, sMon$, sYear$
    sDateStr = Date$
    Select Case Mid$(sDateStr, 1, 2)
        Case "01"
            sMon = "January"
        Case "02"
            sMon = "February"
        Case "03"
            sMon = "March"
        Case "04"
            sMon = "April"
        Case "05"
            sMon = "May"
        Case "06"
            sMon = "June"
        Case "07"
            sMon = "July"
        Case "08"
            sMon = "August"
        Case "09"
            sMon = "September"
        Case "10"
            sMon = "October"
        Case "11"
            sMon = "November"
        Case "12"
            sMon = "December"
    End Select
    If Mid$(sDateStr, 4, 1) = "0" Then
        sDay = Mid$(sDateStr, 5, 1)
    Else
        sDay = Mid$(sDateStr, 4, 2)
    End If
    sYear = Mid$(sDateStr, 7, 4)
    sDateStr = sMon & " " & sDay & ", " & sYear
    DateString = sDateStr
End Function


Sub Main
    Dim CS As Object
    Set CS = CreateObject("Connex.Client")

    If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Or CS.ItemType=19 Then
        Dim nBool%, nInst%, nLine%
        Dim sBorn$, sData$, sDate$, sDied$, sDir$, sHdg$, sInfo$, sTag$, sUsage$
        Dim sMenu() As String
        ' Define menu
        ReDim sMenu(8)
        sMenu(0) = "006/007"
        sMenu(1) = "Campus Wide"
        sMenu(2) = "Law School Community"
        sMenu(3) = "Password Access"
        sMenu(4) = "Free Database"
        sMenu(5) = "Database Format"
        sMenu(6) = "Website Format"
        sMenu(7) = "Ebooks"
        sMenu(8) = "EJournals"
        ' Creates dialog
        Begin Dialog UserDialog 286, 162, "Add Database Notes"
            Text  8, 4, 42, 8, "Choose Field Template:", .Text1
            ListBox  8, 16, 208, 140, sMenu(), .ListBox1
            OKButton  224, 16, 54, 14
            CancelButton  224, 36, 54, 14
        End Dialog
        Dim sDialog as UserDialog
        On Error Resume Next
        Dialog sDialog
        If Err = 102 then
            Exit Sub
        End If

        Select Case sDialog.ListBox1
            Case 0 ' OO6/007
                CS.AddField 99, "006  m     o  d        "
                CS.AddField 99, "007  c ßb r"
            Case 1 ' Campus Wide
                CS.AddField 99, "506  Access restricted to Duke University network, or via library proxy server with valid NetID.ß5 NcD-L"
                CS.AddField 99, "542  ßn Copyright restrictions may apply. See Subscription Agreement for permissible uses."
                CS.AddField 99, "500  Contact Law Library Reference <ref@law.duke.edu> for assistance with this resource. ß5 NcD-L"
            Case 2 ' Law School Only
                CS.AddField 99, "506  Access restricted to Duke Law School cummunity.ß5 NcD-L"
                CS.AddField 99, "542  ßn Copyright restrictions may apply. See Subscription Agreement for permissible uses."
                CS.AddField 99, "500  Contact Law Library Reference <ref@law.duke.edu> for assistance with this resource. ß5 NcD-L"
            Case 3 ' Password Access
                CS.AddField 99, "506  Access via username and password."
                CS.AddField 99, "506  Request username and password from Law Library reference desk. ß5 NcD-L"
                CS.AddField 99, "542  ßn Copyright restrictions may apply. See Subscription Agreement for permissible uses."
                CS.AddField 99, "500  Contact Law Library Reference <ref@law.duke.edu> for assistance with this resource. ß5 NcD-L"
            Case 4 ' Free Database
                CS.AddField 99, "506  Freely available database.ß5 NcD-L"
                CS.AddField 99, "542  ßn Copyright restrictions may apply. See Subscription Agreement for permissible uses."
                CS.AddField 99, "500  Contact Law Library Reference <ref@law.duke.edu> for assistance with this resource. ß5 NcD-L"
            Case 5 ' Database
                CS.AddField 99, "655 7Online databases.ß2 lcgft"
                CS.SetFixedField "Form", "o"
                CS.SetFixedField "SrTp", "d"
            Case 6 ' Websites
                CS.AddField 99, "655 7Websites.ß2 lcgft"
                CS.SetFixedField "Form", "o"
                CS.SetFixedField "SrTp", "w"
            Case 7 ' Ebooks
                CS.AddField 99, "655 7Electronic books.ß2 lcgft"
                CS.SetFixedField "Form", "o"
            Case 8 ' Ejournals
                CS.AddField 99, "655 7Electronic journals.ß2 lcgft"
                CS.SetFixedField "Form", "o"
        End Select

    Else
        MsgBox "An bibliographic record must be displayed in order to use this macro", 0, "Add Database Notes"
    End If
End Sub