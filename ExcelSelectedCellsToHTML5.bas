Sub Export2HTML5()
  hForm.Show
End Sub
Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function
Sub NewHTML5(iTB_mySelection, _
iCB_TableClass, iTB_TableClass, iCB_TableId, iTB_TableId, _
iCB_TDrepTH, iCB_TROdd_Even, _
iCB_TRClass, iTB_TRClass, iCB_TRId, iTB_TRId, _
iCB_TDClass, iTB_TDClass, iCB_TDId, iTB_TDId, _
iCB_Save2File, iTB_Save2File, iCB_Save2FileSc, _
iCB_Save2Cell, iTB_Save2Cell, iCB_Save2CellSc, _
iCB_Copy2CP, iCB_Copy2CPSc)

Range(iTB_mySelection).Select
    iFirstLine = Selection.row
    iFirstCol = Selection.Column
    iLastLine = iFirstLine + Selection.Rows.Count - 1
    iLastCol = iFirstCol + Selection.Columns.Count - 1
    
'sOutput = "<table" & " class=" & Chr(34) & "t-additional" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse; text-align:center; width:80%" & Chr(34)
sOutput = "<table"
If iCB_TableClass = True Then sOutput = sOutput & " class=" & Chr(34) & iTB_TableClass & Chr(34)
If iCB_TableId = True Then sOutput = sOutput & " id=" & Chr(34) & iTB_TableId & Chr(34)
sOutput = sOutput & ">"


r = ""
SpanedCell = ""
For k = iFirstLine To iLastLine
    ce = ""
    If iCB_TRClass = False Then
        r = r & "<tr"
    Else
        r = r & "<tr" & " class=" & Chr(34) & AutoInc(iTB_TRClass, k) & Chr(34)
    End If
    
     If iCB_TRId = True Then
        r = r & " id=" & Chr(34) & AutoInc(iTB_TRId, k) & Chr(34)
    End If
    r = r & ">"

 
    For j = iFirstCol To iLastCol
    If iCB_TDClass = True Then AddClass = "class=" & Chr(34) & AutoInc(iTB_TDClass, k, j) Else AddClass = ""
    If iCB_TROdd_Even = True Then
        If AddClass = "" Then
            AddClass = "class=" & Chr(34)
            If k Mod 2 = 0 Then AddClass = AddClass & "even" Else AddClass = AddClass & "odd" '×¸òíîå / Even
            AddClass = " " & AddClass & Chr(34)
        Else
             If k Mod 2 = 0 Then AddClass = AddClass & " even" Else AddClass = AddClass & " odd" '×¸òíîå / Even
             AddClass = " " & AddClass & Chr(34)
        End If
    End If
    If iCB_TDId = True Then AddId = " id=" & Chr(34) & AutoInc(iTB_TDId, k, j) & Chr(34) Else AddId = ""
        If Cells(k, j) <> "" Then
            If Cells(k, j).MergeArea.Count > 1 Then
                SpanedCell = "<td rowspan=" & Chr(34) & Cells(k, j).MergeArea.Rows.Count & Chr(34) & " colspan=" & Chr(34) & Cells(k, j).MergeArea.Columns.Count & Chr(34) & AddClass & AddId & ">" & Cells(k, j) & "</td>"
                rSpan0 = " rowspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, rSpan0, "")
                
                cSpan0 = " colspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, cSpan0, "")
                ce = ce & SpanedCell
            Else
                ce = ce & "<td" & AddClass & AddId & ">" & Cells(k, j) & "</td>"
            End If
        Else
            If Cells(k, j).MergeArea.Count = 1 Then ce = ce & "<td" & AddClass & AddId & ">&nbsp;</td>"
        End If
    Next j
    If iCB_TDrepTH = True And k = iFirstLine Then ce = Replace(ce, "<td", "<th"): ce = Replace(ce, "</td>", "</th>")
    r = r & ce & "</tr>"
Next k

sOutput = sOutput & r & "</table>"

sOutputDone = sOutput
'---- Ñîõðàíåíèå â ôàéë -----
If iCB_Save2File = True Then
        If iCB_Save2FileSc = True Then sOutputDone = ScreenFunc(sOutputDone)
    cx = SaveTXTfile(iTB_Save2File, sOutputDone)
    If cx = True Then
        callm = MsgBox(Prompt:="Ôàéë " & iTB_Save2File & " ñîõðàí¸í óñïåøíî.", _
        Buttons:=vbOKnly + vbInformation + vbMsgBoxSetForeground, Title:="Ñîõðàíåíèå...")
    Else
        callm = MsgBox(Prompt:="Îøèáêà ïðè ñîõðàíåíèè ôàéëà " & iTB_Save2File & "!!!", _
        Buttons:=vbCritical + vbOKOnly + vbMsgBoxSetForeground, Title:="Ñîõðàíåíèå...")
    End If
End If

sOutputDone = sOutput
'---- Ñîõðàíåíèå â ÿ÷åéêó -----
If iCB_Save2Cell = True Then
    If iCB_Save2CellSc = True Then sOutputDone = ScreenFunc(sOutputDone)
    If Len(sOutputDone) > 32767 Then callm = MsgBox(Prompt:="Èòîãîâûå äàííûå ïðåâûøàþò 32767 ñèìâîëîâ, " & _
        "äîïóñòèìûõ â MS Excel!" & vbNewLine & "Äàííûå áóäóò óñå÷åíû äî 32767 ñèìâîëîâ." & vbNewLine & "(Âñå âîïðîñû â Microsoft)", _
        Buttons:=vbCritical + vbOKOnly + vbMsgBoxSetForeground, Title:="Microsoft Excel...")
    Range(iTB_Save2Cell).Value = sOutputDone
    MsgBox "Äàííûå âûâåäåíû â ÿ÷åéêó " & iTB_Save2Cell
End If

sOutputDone = sOutput
'---- Ñîõðàíåíèå â áóôåð îáìåíà -----
If iCB_Copy2CP = True Then
    If iCB_Copy2CPSc = True Then sOutputDone = ScreenFunc(sOutputDone)
    hForm.TB_Clip = sOutputDone
    hForm.TB_Clip.SelStart = 0
    hForm.TB_Clip.SelLength = hForm.TB_Clip.TextLength
    hForm.TB_Clip.Copy
    MsgBox "Äàííûå ñêîïèðîâàíû â ÁÓÔÅÐ ÎÁÌÅÍÀ." + vbNewLine + "Â íóæíîì ìåñòå íàæìèòå Ctrl+V"
End If
hForm.Hide
End Sub
Private Function AutoInc(ByVal iName As String, Optional ByVal row As Integer, Optional ByVal col As Integer)
    If row <> 0 Then
        row = row - Selection.row + 1
        iName = Replace(iName, "%", row)
    End If

    If col <> 0 Then
        col = col - Selection.Column + 1
        iName = Replace(iName, "$", col)
    End If
    AutoInc = iName
End Function
Function ScreenFunc(ByVal txt As String)
a = Array(92, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 58, 59, 60, 61, 62, 63, 64, 91, 93, 94, 95, 96, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 185, 187)
For Each i In a
    txt = Replace(txt, Chr(i), "\" & Chr(i))
Next
ScreenFunc = txt
End Function
