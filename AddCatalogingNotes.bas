'MacroName: AddCatalogingNotes
'MacroDescription: Adds general RDA style cataloging notes to a record
'
' Written by Sean Chen <schen@law.duke.edu>
' Revised: November 23, 2014
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

Function SortString(InputString$) As String
    Dim i%, Switched%, InputLen%
    Dim FirstChar$, SecondChar$
    If InputString$ <> "" Then
        InputLen% = Len( InputString$ )
        If InputLen% > 0 Then
            Switched = TRUE
            Do
              For i = 1 To InputLen% - 1
                FirstChar$  = Mid$( InputString$, i, 1 )
                SecondChar$ = Mid$( InputString$, i + 1, 1 )
                If FirstChar$ > SecondChar$ Then
                    Mid$( InputString$, i, 1 ) = SecondChar$
                    Mid$( InputString$, i  + 1, 1 ) = FirstChar$
                    Switched = TRUE
                  Else
                    Switched = FALSE
                End If
              Next i
            Loop Until ( Switched = FALSE And i = InputLen% )
        End If
    SortString = InputString
    End If
End Function

Function GetUpdatedFixedField(sFieldName$, sDefault$, sValue$)
    Dim CS As Object
    Set CS = CreateObject("Connex.Client")

    If CS.GetFixedField(sFieldName, sValue) = False Then
        sValue = sDefault
    End If
    If InStr(sValue, sDefault) = 0 Then
        sValue = sValue & sDefault
    End If
    sValue = SortString(sValue)
    GetUpdatedFixedField = sValue
End Function

Sub Main
    Dim CS As Object
    Set CS = CreateObject("Connex.Client")

    If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Or CS.ItemType=19 Then
        Dim nBool%, nInst%, nLine%, bUpdateCont%, bUpdateIndx%
        Dim sBorn$, sData$, sDate$, sDied$, sDir$, sHdg$, sInfo$, sTag$, sUsage$
        Dim sNote$, sMatch$, sYear$
        Dim sCont$
        Dim sMenu() As String
        ' Define menu
        ReDim sMenu(12)
        sMenu(0) = "504 Includes bibliographical references."
        sMenu(1) = "504 Includes bibliographical references (pages <page-span>)."
        sMenu(2) = "504 Includes bibliographical references and index."
        sMenu(3) = "504 Includes bibliographical references (pages <page-span>) and index."
        sMenu(4) = "500 Include index."
        sMenu(5) = "502 Dissertation (RDA 7.9.1.3)"
        sMenu(6) = "500 Originally presented as the author's (RDA 7.9.1.3)"
        sMenu(7) = "588 Title from cover (RDA 2.17.2)"
        sMenu(8) = "500 Revision of the author's <Title of original work> (Unstructured)"
        sMenu(9) = "775 $i revision of: $a <Title of original work> "
        sMenu(10) = "588 DBO online resource; title from PDF title page (<Provider>, viewed <Today>)'"
        sMenu(11) = "588 DBO print version record."

        ' Creates dialog
        Begin Dialog UserDialog 326, 202, "Add Cataloging Notes"
            Text  8, 4, 42, 8, "Choose Field Template:", .Text1
            ListBox  8, 16, 248, 140, sMenu(), .ListBox1
            OKButton  264, 16, 54, 14
            CancelButton  264, 36, 54, 14
        End Dialog
        Dim sDialog as UserDialog
        On Error Resume Next
        Dialog sDialog
        If Err = 102 then
            Exit Sub
        End If

        sYear = Mid$(Date$, 7, 4)
        sMatch = "ü"               
        Select Case sDialog.ListBox1
            Case 0 ' 504 Includes bibliographical references.
                sNote = "504  Includes bibliographical references."
                sCont = GetUpdatedFixedField("Cont", "b", sCont)
            Case 1 ' 504 Includes bibliographical references (pages <page-span>)
                sNote = "504  Includes bibliographical references (pages <page-span>)."
                sMatch = "(pages <page-span>)"
                sCont = GetUpdatedFixedField("Cont", "b", sCont)                
            Case 2 ' 504 Includes bibliographical references and index
                sNote = "504  Includes bibliographical references and index."
                sCont = GetUpdatedFixedField("Cont", "b", sCont)
                If CS.SetFixedField("Indx", "1") = True Then
                    bUpdateIndx = True
                End If
            Case 3 ' 504 Includes bibliographical references (pages <page-span>) and index
                sNote = "504  Includes bibliographical references (pages <page-span>) and index."
                sMatch = "(pages <page-span>)"
                sCont = GetUpdatedFixedField("Cont", "b", sCont)
                If CS.SetFixedField("Indx", "1") = True Then
                    bUpdateIndx = True
                End If
            Case 4 ' 500 Includes index
                sNote = "500  Includes index."
                If CS.SetFixedField("Indx", "1") = True Then
                    bUpdateIndx = True
                End If
            Case 5 ' Thesis note
                sNote = "502  ßb Ph. D. ßc <Institution heading> ßd " & sYear & "."
                sMatch = "<Institution heading>"
                sCont = GetUpdatedFixedField("Cont", "m", sCont)
            Case 6 ' Originally presented
                sNote = "500  Originally presented as the author's thesis (doctoral)--<Institution heading>, " & sYear & "."
                sMatch = "<Institution heading>"
            Case 7 ' Title from cover
                sNote = "588  Title from cover."
            Case 8 ' Revision of the author's <Title of original work>
                sNote = "500  Revision of the author's <Title of original work>"
                sMatch = "<Title of original work>"
            Case 9 '775 $i revision of: $a <Title of original work>
                sNote = "77508<OCLC Number> "
                sMatch = "<OCLC Number>"
                MsgBox("Insert OCLC Number and Insert From Cited Record and then add $i Revision of: ")
            Case 10 '588 Description based on online resource; title from  PDF title page (<Provider>, viewed Date'
                sNote = "588  Description based on online resource; title from PDF title page (<Provider>, viewed "
                sNote = sNote & DateString & ")."
                sMatch = "<Provider>"
            Case 11 '588 Description based on print version record.
                sNote = "588  Description based on print version record."
                
        End Select
        If sCont <> "" and CS.SetFixedField("Cont", sCont) = True Then
            bUpdateCont = True
        End If
        If CS.AddField (99, sNote) = True Then
            CS.CursorColumn = 6
            CS.FindText sMatch, False
        End If        

    Else
        MsgBox "An bibliographic record must be displayed in order to use this macro", 0, "Add Cataloging Notes"
    End If
End Sub