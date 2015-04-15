Attribute VB_Name = "Excel2HTML"

Sub NewHTML2()
Attribute NewHTML2.VB_ProcData.VB_Invoke_Func = "ü\n14"
 Selection.Areas(1).Select    ' íà ñëó÷àé âûäåëåíèÿ íåñâÿçàííûõ äèàïàçîíîâ

    iFirstLine = Selection.row
    iFirstCol = Selection.Column
    iLastLine = iFirstLine + Selection.Rows.Count - 1
    iLastCol = iFirstCol + Selection.Columns.Count - 1
    
sOutput = "<table align=" & Chr(34) & "left" & Chr(34) & " class=" & Chr(34) & "t-additional" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse; text-align:center; width:80%" & Chr(34)
r = ""
SpanedCell = ""
For k = iFirstLine To iLastLine
    ce = ""
    r = r & "<tr>"
    For j = iFirstCol To iLastCol
        If Cells(k, j) <> "" Then
            If Cells(k, j).MergeArea.Count > 1 Then
                SpanedCell = "<td rowspan=" & Chr(34) & Cells(k, j).MergeArea.Rows.Count & Chr(34) & " colspan=" & Chr(34) & Cells(k, j).MergeArea.Columns.Count & Chr(34) & ">" & Cells(k, j) & "</td>"
                rSpan0 = " rowspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, rSpan0, "")
                
                cSpan0 = " colspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, cSpan0, "")
                ce = ce & SpanedCell
            Else
                ce = ce & "<td>" & Cells(k, j) & "</td>"
            End If
        Else
            If Cells(k, j).MergeArea.Count = 1 Then ce = ce & "<td>&nbsp;</td>"
        End If
    Next j
    If k = iFirstLine Then ce = Replace(ce, "td", "th")
    r = r & ce & "</tr>"
Next k

sOutput = sOutput & r & "</table>"
cx = SaveTXTfile("C:\Users\Dr\Desktop\Opera temp\___NEW\" & ActiveSheet.Name & ".html", sOutput)
If cx = True Then
    MsgBox "Ôàéë ñîõðàí¸í óñïåøíî."
Else
    MsgBox "Îøèáêà ïðè ñîõðàíåíèè ôàéëà!" 'vbCritical+vbOKOnly
End If
End Sub

Public Sub SetTextIntoClipboard(ByVal txt As String)
    Dim MyDataObj As New DataObject
    MyDataObj.SetText txt
    MyDataObj.PutInClipboard
End Sub

Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function
Sub NewHTML3()
 Selection.Areas(1).Select    ' íà ñëó÷àé âûäåëåíèÿ íåñâÿçàííûõ äèàïàçîíîâ

    iFirstLine = Selection.row
    iFirstCol = Selection.Column
    iLastLine = iFirstLine + Selection.Rows.Count - 1
    iLastCol = iFirstCol + Selection.Columns.Count - 1
    
sOutput = "<table align=" & Chr(34) & "left" & Chr(34) & " class=" & Chr(34) & "t-additional" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse; text-align:center; width:80%" & Chr(34)
r = ""
SpanedCell = ""
For k = iFirstLine To iLastLine
    ce = ""
    r = r & "<tr>"
    For j = iFirstCol To iLastCol
        If Cells(k, j) <> "" Then
            If Cells(k, j).MergeArea.Count > 1 Then
                SpanedCell = "<td rowspan=" & Chr(34) & Cells(k, j).MergeArea.Rows.Count & Chr(34) & " colspan=" & Chr(34) & Cells(k, j).MergeArea.Columns.Count & Chr(34) & ">" & Cells(k, j) & "</td>"
                rSpan0 = " rowspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, rSpan0, "")
                
                cSpan0 = " colspan=" & Chr(34) & 1 & Chr(34)
                SpanedCell = Replace(SpanedCell, cSpan0, "")
                ce = ce & SpanedCell
            Else
                ce = ce & "<td>" & Cells(k, j) & "</td>"
            End If
        Else
            If Cells(k, j).MergeArea.Count = 1 Then ce = ce & "<td>&nbsp;</td>"
        End If
    Next j
    If k = iFirstLine Then ce = Replace(ce, "td", "th")
    r = r & ce & "</tr>"
Next k

sOutput = sOutput & r & "</table>"
cx = SaveTXTfile("C:\Users\Dr\Desktop\Opera temp\___NEW\" & ActiveSheet.Name & ".html", sOutput)
If cx = True Then
    MsgBox "Ôàéë ñîõðàí¸í óñïåøíî."
Else
    MsgBox "Îøèáêà ïðè ñîõðàíåíèè ôàéëà!" 'vbCritical+vbOKOnly
End If
End Sub
