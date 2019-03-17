Attribute VB_Name = "M_Bkm"
Type TfntParam
    name As String
    val As String
    lengt As Integer
End Type

Type TpartStr
    str As String
    isParam As Boolean
    length As Integer
End Type


Function setTextBkm(ByVal bkmName As String, ByVal bkmText As String) As Boolean
Dim oRng As Range
    
    If ActiveDocument.Bookmarks.Exists(bkmName) = True Then
        Set oRng = ActiveDocument.Bookmarks(bkmName).Range
        oRng.text = bkmText
        'Re-insert the bookmark
        ActiveDocument.Bookmarks.Add bkmName, oRng
        setTextBkm = True
    Else
        Debug.Print "M_Bkm => setTextBkm => Закладка с именем " & bkmName & " отсуствует"
        setTextBkm = False
    End If
End Function


Function setTextWithFontsBkm(bkmName As String, bkmText As String) As Boolean
Dim oRng As Range, oRngPart As Range
Dim fntTxt As String
Dim start%, startPos%, endPos%
Dim partStr As TpartStr

    If ActiveDocument.Bookmarks.Exists(bkmName) = True Then
        Set oRng = ActiveDocument.Bookmarks(bkmName).Range
        oRng.text = ""
        
        Set oRngPart = oRng
        
        start = 1
        endPos = oRng.start
        Do While start < Len(bkmText)
            partStr = getPart(bkmText, start)
            Debug.Print "M_Bkm = > setTextWithFontsBkm => partStr.str=", partStr.str
            start = start + partStr.length
            
            If (partStr.isParam = True) Then
                setFont oRngPart, partStr.str
            Else
                startPos = endPos
                endPos = startPos + Len(partStr.str)
                oRng.InsertAfter text:=partStr.str
                Set oRngPart = ActiveDocument.Range(start:=startPos, End:=endPos)
            End If
        Loop
        
        ActiveDocument.Bookmarks.Add bkmName, oRng
        setTextWithFontsBkm = True
    Else
        Debug.Print "M_Bkm => setTextWithFontsBkm => Закладка с именем " & bkmName & " отсуствует"
        setTextWithFontsBkm = False
    End If
    
End Function

' если длина текста в закладке больше length и закладка в ячейке таблицы
' то зжимаем по содержимому и убираем переном
' иначе наоборот
Function fitBkmInCell(bkmName As String, length As Integer) As Boolean
Dim oRng As Range
Dim oCell As Cell
    
    If (ActiveDocument.Bookmarks.Exists(bkmName)) Then
        Set oRng = ActiveDocument.Bookmarks(bkmName).Range
        If (oRng.Cells.Count > 0) Then
            Set oCell = oRng.Cells(1)
        Else
            Debug.Print "M_Bkm => fitBkmInCell  Cells.Count = 0"
            fitBkmInCell = False
            Exit Function
        End If
        With oCell
            If (Len(oRng.text) > length) Then
                .WordWrap = False
                .FitText = True
            Else
                .WordWrap = True
                .FitText = False
            End If
        End With
        fitBkmInCell = True
        Exit Function
    Else
        fitBkmInCell = False
        Exit Function
    End If
End Function


Private Function parseParam(str As String, Optional start% = 1, Optional delimeter$ = ":") As TfntParam
Dim fntParam As TfntParam
Dim n%
    n = InStr(start, str, delimeter)
    If (n > 0) Then
        fntParam.name = Mid(str, 1, n - 1)
        fntParam.val = Mid(str, n + 1, Len(str) - n)
    Else
        fntParam.name = ""
        fntParam.val = ""
    End If
    fntParam.lengt = Len(str)
    parseParam = fntParam
End Function


Private Function getParamFromStr(str As String, Optional start% = 1, Optional delimeter$ = ";") As TfntParam
Dim simb As String
Dim prevSimb As String
Dim strLen As Integer
Dim subStr As String
Dim fntParam As TfntParam
    
    strLen = Len(str)
    If (start >= strLen) Then
        Exit Function
    End If
    
    prevSimb = ""
    For i = start To strLen
        simb = Mid(str, i, 1)
        If (simb = delimeter And prevSimb <> "\") Then
            subStr = Mid(str, start, i - start)
            getParamFromStr = parseParam(subStr)
            Exit Function
        End If
        prevSimb = simb
    Next i
    subStr = Mid(str, start, strLen - start + 1)
    getParamFromStr = parseParam(subStr)
End Function


Private Function getPart(str$, Optional start% = 1) As TpartStr
Dim partStr As TpartStr
Dim fstSmb As String
Dim n%, length%
    
    fstSmb = Mid(str, start, 1)
    'Debug.Print "getPart fstSmb: ", fstSmb
    If (fstSmb = "{") Then
        ' если первый символ "{" - открывабщая скобка параметров
        ' ищем закрывающую скобку "}"
        n = InStr(start + 1, str, "}")
        If (n > 0) Then
            partStr.length = n - start + 1 ' длина с учетом скобок
            partStr.str = Mid(str, start, partStr.length)
            partStr.isParam = True
        Else
            partStr.length = Len(str) - start
            partStr.str = Mid(str, start)
            partStr.isParam = False
        End If
        getPart = partStr
        Exit Function
    Else
        ' если первый символ НЕ "{"
        ' ищем открывающую скобку парамтеров "{"
        n = InStr(start + 1, str, "{")
        If (n > 0) Then
            partStr.length = n - start
            partStr.str = Mid(str, start, partStr.length)
            partStr.isParam = False
        Else
            partStr.length = Len(str) - start
            partStr.str = Mid(str, start)
            partStr.isParam = False
        End If
        getPart = partStr
        Exit Function
    End If
End Function


Private Sub setFont(oRng As Range, fntParamsStr As String)
Dim fntParam As TfntParam
Dim fntParamsAr() As String
Dim oFnt As font
Dim start%
    
    'удяляем скобки "{....}"
    fntParamsStr = Mid(fntParamsStr, 2, Len(fntParamsStr) - 2)
    
    Set oFnt = oRng.font
    start = 1
    
    fntParamsAr = Split(fntParamsStr, ";")
    
    
    Do While start < Len(fntParamsStr)
        fntParam = getParamFromStr(str:=fntParamsStr, start:=start)
        'Debug.Print "M_Bkm => setFont => fntParamsStr= ", fntParamsStr, "", fntParam.name, ": ", fntParam.val
        Select Case fntParam.name
            Case "Name"
                oRng.font.name = fntParam.val
            Case "Size"
                oRng.font.Size = CInt(fntParam.val)
            Case "Bold"
                oRng.font.Bold = CBool(fntParam.val)
            Case Else
                '
        End Select
        start = start + fntParam.lengt + 1
    Loop
End Sub


Sub test_setFont()
Dim bkmName As String, bkmText As String

    bkmName = "testBkm"
    bkmText = "Раздел №8.{Name:Times New Roman;Bold:True;Size:14} Мероприятия по обеспечению пожарной безопасности для{Bold:False}"
    
    setTextWithFontsBkm bkmName, bkmText
End Sub



Private Sub dsf()
    'ActiveDocument.Range(start:=50, End:=50).InsertParagraph
    
    Dim SrtTxtGIP
    SrtTxtGIP = "А.П. Вываыывалдывта"
    Debug.Print Right(SrtTxtGIP, Len(SrtTxtGIP) - InStrRev(SrtTxtGIP, " "))  ' обрезаем инициалы слева
End Sub


