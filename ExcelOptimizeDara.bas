'Option Base 1
Option Explicit
Dim MainArray(), DEArray()
Dim MyTMPSheetName, MySheetName, TableVPRName As String
Dim k, VeGran As Byte
Sub MyNewMagic()
  MForm.Show 0
End Sub
Sub î÷èñòèòü(Optional FakeArgument As Integer)
Unload MForm
MForm.Show 0
End Sub
Sub NewMagicExecutor(Optional FakeArgument As Integer)
Call GeneralModule.TurboON
Call ReadMyVal
If k > 0 Then Call ProcessMyData(MainArray, DEArray) Else Call ProcessMyData(MainArray)
    Sheets(MySheetName).Activate
Call FormatThis(VeGran)
Call CalcThis
MForm.Hide
Sheets(MyTMPSheetName).Delete
Call GeneralModule.TurboOF
Selection.Activate
End Sub
Sub ERRORCOLLECTOR(Optional FakeArgument As Integer)
Dim MyPrompt As String
Dim MyMessage As Variant
If MForm.tUKTZED = False Then
    MyPrompt = "Áåç êîäà ÓÊÒ ÇÅÄ íåò ñìûñëà â äàëüíåéøèõ äåéñòâèÿõ." & vbLf & "Óêàæèòå êîä ÓÊÒ ÇÅÄ"
    MyMessage = MsgBox(Prompt:=MyPrompt, Buttons:=vbExclamation, Title:="×òî-òî ïîøëî íå òàê..")
    Exit Sub
ElseIf MForm.tBRUT = False Then
    MyPrompt = "Áåç âåñà ÁÐÓÒÒÎ íåò ñìûñëà â äàëüíåéøèõ äåéñòâèÿõ." & vbLf & "Óêàæèòå âåñ ÁÐÓÒÒÎ"
    MyMessage = MsgBox(Prompt:=MyPrompt, Buttons:=vbExclamation, Title:="×òî-òî ïîøëî íå òàê..")
    Exit Sub
ElseIf MForm.tPP = False Then
    MyPrompt = "Íå óêàçàíî ïîëå ¹ ï/ï" & _
        vbLf & "Áåç íåãî áóäåò òðóäíî ÷òî-òî íàéòè" & _
        vbLf & vbLf & "ÏÐÎÑÒÀÂÈÒÜ ''¹ï/ï'' ÀÂÒÎÌÀÒÈ×ÅÑÊÈ?"
    MyMessage = MsgBox(Prompt:=MyPrompt, Buttons:=vbExclamation + vbYesNo, Title:="×òî-òî ïîøëî íå òàê..")
    If MyMessage = vbYes Then Call AutoNumeration: Call ERRORCOLLECTOR
Else
    Call NewMagicExecutor
End If
End Sub
Private Sub AutoNumeration()
Dim LC, LR, myI As Long
LC = Cells(1, Columns.Count).End(xlToLeft).Column
LR = Cells(Rows.Count, Range(MForm.oUKTZED.Value).Column).End(xlUp).row
Cells(1, LC + 1).Value = "¹ï/ï"
For myI = 2 To LR
    Cells(myI, LC + 1).Value = myI - 1
Next myI
Cells(1, LC + 1).Select
MForm.tPP = True
Call MForm.tPP_Click
End Sub
Private Sub ReadMyVal()
Dim AddArray()
Dim iControl, jControl, kControl As Control
Dim m, i, j, ArtIsHere, oI, oI2, oI3, oIFIX, myPos, MyPlace As Byte
Dim LR, LC, xI, yI As Long
Dim iFind As Boolean
Dim ZASRANETS As String

LR = Cells(Rows.Count, Range(MForm.oUKTZED.Value).Column).End(xlUp).row

m = 0
ArtIsHere = 0
j = 0
k = 0

For Each iControl In MForm.Controls ' - ñ÷èòàåì ñêîëüêî íàæàòî îñíîâíûõ
    If Mid(iControl.Name, 1, 1) = "t" Then
        If iControl = True Then m = m + 1
    End If
Next


If MForm.tARTXY = True Then ' - ñ÷èòàåì, ñêîëüêî íàæàòî äîïîëíèòåëüíûõ
    ArtIsHere = 1
    For Each jControl In MForm.Controls
        If Mid(jControl.Name, 1, 1) = "b" And jControl.Visible = True Then
            If jControl = True Then
                j = j + 1
                If MForm.Controls("d" & Mid(jControl.Name, 2, 1)) = True Then k = k + 1: ZASRANETS = MForm.Controls("d" & Mid(jControl.Name, 2, 1)).Name
            End If
        End If
    Next
End If
If j = 0 Then MForm.tARTXY = False: Call MForm.ClearAC

If k = 1 Then k = 0: MForm.Controls(ZASRANETS) = False '- ëîâèì, åñëè îáúåäèíÿòü òîëüêî îäèí êðèòåðèé, òî íå íàäî

If k > 0 Then '== çàãîíÿåì, äîï. êðèòåðèè, êîò. íàäî îáúåäåíèòü - â îäèí îäíîìåð. ìàññèâ
    xI = 0
    For oI3 = 1 To j
        For Each jControl In MForm.Controls
           If Mid(jControl.Name, 1, 1) = "d" Then
               If Val(MForm.Controls("a" & Mid(jControl.Name, 2, 1))) = oI3 Then
                   If jControl = True And jControl.Visible = True Then
                       ReDim Preserve AddArray(xI)
                       AddArray(xI) = Range(MForm.Controls("b" & Mid(jControl.Name, 2, 1)).Caption).Column
                       xI = xI + 1
                   End If
               End If
           End If
        Next
    Next oI3

    ReDim DEArray(LR)
    For xI = 1 To LR
        For yI = 0 To UBound(AddArray)
            DEArray(xI) = DEArray(xI) & ActiveSheet.Cells(xI, AddArray(yI)).Value & MForm.vBox1.Value
        Next yI
        DEArray(xI) = Mid(DEArray(xI), 1, Len(DEArray(xI)) - Len(MForm.vBox1.Value))
    Next xI
     
'    For xI = 1 To UBound(DEArray) '- ÷òîáû âûâåñòè è ïîñìîòðåòü
'     Cells(xI, 1).Value = DEArray(xI)
'    Next xI
End If


i = m - ArtIsHere + j - k

If k > 0 Then i = i + 1
ReDim MainArray(1 To 2, 1 To i)
oI = 1
If MForm.tPP = True Then
    MainArray(1, oI) = MForm.tPP.Caption
    MainArray(2, oI) = Range(MForm.oPP.Value).Column
    oI = oI + 1
End If

If MForm.tUKTZED = True Then
    MainArray(1, oI) = MForm.tUKTZED.Caption
    MainArray(2, oI) = Range(MForm.oUKTZED.Value).Column
    oI = oI + 1
End If

If MForm.tNAME = True Then
    MainArray(1, oI) = MForm.tNAME.Caption
    MainArray(2, oI) = Range(MForm.oNAME.Value).Column
    oI = oI + 1
End If
oIFIX = oI

iFind = False
If MForm.tARTXY.Visible = True And MForm.tARTXY = True Then
    For oI2 = 1 To j
        For Each kControl In MForm.Controls
            If Mid(kControl.Name, 1, 1) = "b" Then
                If kControl = True And kControl.Visible = True Then
                    If MForm.Controls("d" & Mid(kControl.Name, 2, 1)) = False Then
                        If Val(MForm.Controls("a" & Mid(kControl.Name, 2, 1)).Caption) = oI2 Then
                            MainArray(1, oI + Val(MForm.sARTXY.Caption) - 1) = MForm.Controls("c" & Mid(kControl.Name, 2, 1)).Value
                            MainArray(2, oI + Val(MForm.sARTXY.Caption) - 1) = Range(kControl.Caption).Column
                            oI = oI + 1
                        End If
                    Else
                        If k > 0 And iFind = False And Val(MForm.Controls("a" & Mid(kControl.Name, 2, 1)).Caption) = oI2 Then
                            MainArray(1, oI + Val(MForm.sARTXY.Caption) - 1) = DEArray(1)
                            MainArray(2, oI + Val(MForm.sARTXY.Caption) - 1) = "íåò!!!"
                            iFind = True
                            oI = oI + 1
                        End If
                    End If
                End If
            End If
        Next
    Next oI2
End If

For Each iControl In MForm.Controls
    If iControl.Name = "tQTYXY" Or iControl.Name = "tEDEXY" Or iControl.Name = "tNETXY" Then
        If iControl = True Then
            myPos = Val(MForm.Controls("s" & Mid(iControl.Name, 2, 5)).Caption)
            If MForm.tARTXY = True And myPos > Val(MForm.sARTXY.Caption) Then
                MyPlace = oI - 1
            Else
                MyPlace = oIFIX - 1  ' - òóò äîáàâèë -1
            End If
            MainArray(1, MyPlace + myPos - 1) = iControl.Caption
            MainArray(2, MyPlace + myPos - 1) = Range(MForm.Controls("o" & Mid(iControl.Name, 2, 5)).Value).Column
        End If
    End If
Next


If MForm.tBRUT = True Then
    MainArray(1, i) = MForm.tBRUT.Caption
    MainArray(2, i) = Range(MForm.oBRUT.Value).Column
End If

End Sub

Private Sub ProcessMyData(MainArray As Variant, Optional DEArray As Variant)
Dim i, j, LR, LC As Long

ActiveSheet.Copy After:=ActiveSheet
MyTMPSheetName = ActiveSheet.Name
ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
Sheets.Add After:=Sheets(Sheets.Count)
MySheetName = "X_" & Replace(Time, ":", "") & Left(Replace(Date, ".", ""), 4) & Right(Year(Date), 2)
ActiveSheet.Name = MySheetName
Sheets(MySheetName).Tab.Color = 5287936
Sheets(MyTMPSheetName).Activate

LR = Cells(Rows.Count, Range(MForm.oUKTZED.Value).Column).End(xlUp).row
LC = Cells(1, Columns.Count).End(xlToLeft).Column
VeGran = (UBound(MainArray, 1) - LBound(MainArray, 1) + 1) * (UBound(MainArray, 2) - LBound(MainArray, 2) + 1) / 2

MForm.GGGGSUM.Caption = Application.WorksheetFunction.Sum(Range(MForm.oBRUT.Value, Cells(LR, Range(MForm.oBRUT.Value).Column)))
For j = 1 To VeGran
    Sheets(MySheetName).Cells(1, j) = MainArray(1, j)
    For i = 2 To LR
        If MainArray(2, j) = "íåò!!!" Then
            Sheets(MySheetName).Cells(i, j).Value = DEArray(i)
        Else
            Sheets(MySheetName).Cells(i, j).Value = Sheets(MyTMPSheetName).Cells(i, Val(MainArray(2, j))).Value
        End If
    Next i
Next j

End Sub

Private Sub FormatThis(VeGran As Byte)
Dim oblast, allregion As String
Dim LR, LC, Q As Long
Dim w As Integer

LR = Cells(Rows.Count, Range("B1").Column).End(xlUp).row
LC = Cells(1, Columns.Count).End(xlToLeft).Column
oblast = Range("B2", Cells(LR, 2)).Address '"J2:J40"
allregion = Range(Cells(1, 1), Cells(LR, LC)).Address

ActiveSheet.Sort.SortFields.Clear
ActiveSheet.Sort.SortFields.Add Key:=Range( _
        oblast), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(allregion)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Call GeneralModule.TurboOF
 Cells(2, 2).Select
 Call MyMacro.ÐÀÇÓÊÐÀØÊÀ
Call GeneralModule.TurboON
 
With Range("A1", Cells(1, VeGran)).Font '//ØÐÈÔÒ ØÀÏÊÈ//
   .Color = RGB(255, 255, 255)
   .TintAndShade = 0
   .Name = "Arial"
   .Size = 12
   .OutlineFont = True
   .Underline = xlUnderlineStyleNone
   .Bold = True
End With

With Range("A1", Cells(1, VeGran)) '//ÎÔÎÐÌËÅÍÈÅ ØÀÏÊÈ//
   .Interior.Color = RGB(0, 0, 0)
   .HorizontalAlignment = xlCenter
   .VerticalAlignment = xlCenter
   .WrapText = False
   .Orientation = 0
   .MergeCells = False
   .ShrinkToFit = False
End With

With Range(Cells(2, 1), Cells(LR, LC))
   .Font.Name = "Arial"
   .Font.Color = RGB(0, 0, 0)
   .Font.Size = 10
   .WrapText = False
   .HorizontalAlignment = xlCenter
   .VerticalAlignment = xlCenter
   .MergeCells = False
   .NumberFormat = "General"
End With

Range("A1").ColumnWidth = 5
Range("B1").ColumnWidth = 12
For Q = 4 To LC - 1
    w = Int(Val(Len(Cells(1, Q).Text))) + 3
    Cells(1, Q).Select
    Selection.ColumnWidth = w
Next Q

With Range(Cells(2, LC), Cells(LR, LC))
   .ColumnWidth = 14
   .NumberFormat = "#,##0.000"
End With

If MForm.tNAME = True Then
    With Range(Cells(2, 3), Cells(LR, 3))
        .HorizontalAlignment = xlLeft
        .ColumnWidth = 28
    End With
End If
 
End Sub

Private Sub CalcThis()
Dim LR, LC, LRKrit, G As Long
Dim TableVPR, DISU As String
Dim SeResult As Variant
Dim SumBrutTotal As Double
LR = Cells(Rows.Count, Range("B1").Column).End(xlUp).row
LC = Cells(1, Columns.Count).End(xlToLeft).Column
Columns(2).Select
Selection.Copy
Columns(LC + 3).Select
ActiveSheet.Paste
Application.CutCopyMode = False

    ActiveSheet.Range(Cells(2, LC + 3), Cells(LR, LC + 3)).RemoveDuplicates Columns:=1, Header:=xlNo
    TableVPR = "=R2C2:R" & LR & "C" & LC
    TableVPRName = "ÒàáëÂÏÐ" & Replace(Time, ":", "") & Left(Replace(Date, ".", ""), 4) & Right(Year(Date), 2)
    ActiveWorkbook.Names.Add Name:=TableVPRName, RefersToR1C1:=TableVPR
    Selection.Columns.AutoFit
    
If MForm.tNAME = True Then
    Call myVLOOKUP(MForm.tNAME, 4)
End If
    
If MForm.tEDEXY = True Then
   Call myVLOOKUP(MForm.tEDEXY, 6)
End If
    
    LRKrit = Cells(Rows.Count, LC + 3).End(xlUp).row
    Cells(1, LC + 2).Value = "# ï/ï"
    For k = 2 To LRKrit
        Cells(k, LC + 2).Value = k - 1
    Next k
    
If MForm.tQTYXY = True Then
    Call mySUMIF(MForm.tQTYXY, 5)
End If

G = 7
If MForm.tNETXY = True Then
    Call mySUMIF(MForm.tNETXY, G)
    G = 8
End If

If MForm.tBRUT = True Then
    Call mySUMIF(MForm.tBRUT, G)
End If
        
'==============ÂÑÅ ÔÎÐÌÓËÛ (è ÂÏÐ, è ÑÓÌÌ) íàïèñàíû ïî îäíîé - ÐÀÇÍÎÑÈÌ èõ íà îñòàëüí. ÿ÷åéêè =======
    Range(Cells(2, LC + 4), Cells(2, LC + G)).Select
    Selection.AutoFill Destination:=Range(Cells(2, LC + 4), Cells(LRKrit, LC + G)), Type:=xlFillDefault
'=====================================================================================================
    
    DISU = Range(Cells(2, LC + G), Cells(LRKrit, LC + G)).Address(False, False)
    Cells(LRKrit + 2, LC + G).FormulaLocal = "=ÑÓÌÌ(" & DISU & ")"
    Cells(LRKrit + 2, LC + G).Font.Bold = True
    
    Cells(LRKrit + 2, LC + 5).Value = Val(MForm.GGGGSUM.Caption)
    
    Cells(LRKrit + 2, LC + 6).Value = "Ðàçíèöà:"
    Cells(LRKrit + 3, LC + 6).FormulaLocal = _
    "=" & Cells(LRKrit + 2, LC + G).Address(False, False) & "-" & Cells(LRKrit + 2, LC + 5).Address(False, False)
    
    Range(Cells(2, LC + 7), Cells(LRKrit + 2, LC + G)).Select
    Call MyMacro.Ïðàâ_ôîðìàò_è_âûðàâí
    ActiveSheet.Columns(2).AutoFit
Range(Cells(LRKrit + 2, LC + G), Cells(LRKrit + 2, LC + G)).Select
End Sub

Private Sub myVLOOKUP(iControl As Control, PL As Long)
Dim w, SeResult As Long
Dim Ran As String
For w = 1 To VeGran
        If MainArray(1, w) = iControl.Caption Then
            SeResult = w
            Exit For
        End If
Next w
    Ran = Cells(2, Columns(VeGran + 3).Column).AddressLocal(False, False, xlA1)
    Cells(2, (VeGran + PL)).Formula = "=VLOOKUP(" & Ran & "," & TableVPRName & "," & Columns(SeResult).Column - 1 & ",0)"
    Cells(1, SeResult).Copy: Cells(1, (VeGran + PL)).Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
End Sub
Private Sub mySUMIF(iControl As Control, PL As Long)
Dim DIAP, KRIT, DISU As String
Dim w, LR, LC, LRKrit, SeResult As Long

LR = Cells(Rows.Count, Range("B1").Column).End(xlUp).row


LRKrit = Cells(Rows.Count, VeGran + 3).End(xlUp).row

DIAP = Range(Cells(2, 2), Cells(LR, 2)).Address(False, False)
KRIT = Range(Cells(2, VeGran + 3), Cells(LRKrit, VeGran + 3)).Address(False, False)
    
For w = 1 To VeGran
    If MainArray(1, w) = iControl.Caption Then
        SeResult = w
        Exit For
    End If
Next w
                
DISU = Range(Cells(2, SeResult), Cells(LR, SeResult)).Address(False, False)

     '-âñòàâëÿåì ïåðåìåííûå â ôîðìóëó
    Cells(2, VeGran + PL).FormulaLocal = "=ÑÓÌÌ(ÑÓÌÌÅÑËÈ(" & DIAP & ";" & KRIT & ";" & DISU & "))"
    Cells(1, SeResult).Copy: Cells(1, (VeGran + PL)).Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
    
   
End Sub
