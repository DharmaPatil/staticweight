VERSION 5.00
Begin VB.UserControl GdhPrintWeight 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GdhPrintWeight.ctx":0000
   ScaleHeight     =   570
   ScaleWidth      =   555
End
Attribute VB_Name = "GdhPrintWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'制表符
Private Const strLeftTop As String = "┌"
Private Const strCenterTop As String = "┬"
Private Const strRightTop As String = "┐"
Private Const strLeftCenter As String = "├"
Private Const strCenter As String = "┼"
Private Const strRightCenter As String = "┤"
Private Const strLeftButtom As String = "└"
Private Const strCenterButtom As String = "┴"
Private Const strRightButtom As String = "┘"
Private Const strHLine As String = "─"
Private Const strVLine As String = "│"


Private Enum AlignEnum
    emeLeft = 0
    emeRight = 1
    emeCenter = 2
End Enum

Private Enum PrintStyle
    emeRow = 0
    emePage = 1
End Enum

Private m_Tital As String
Private m_Style As Integer
Private m_Corp As String

Private Function StringChange(vString As String)
    If vString = "0.000" Then
        StringChange = ""
    Else
        StringChange = vString
    End If
End Function

Private Function StringConnection(vString As String, vNum As Integer)
    Dim i As Integer
    Dim retString As String
    
    retString = ""
    For i = 1 To vNum
        retString = retString + vString
    Next i
    StringConnection = retString
End Function

Private Function StringFormat(strSource As String, iLength As Integer, blnDirection As Boolean)
    Dim iStringLen As Integer
    Dim iSpaceNo As Integer
    Dim i As Integer
    
    iStringLen = Len(strSource)
    If iLength > iStringLen Then
        StringFormat = strSource
        iSpaceNo = iLength - iStringLen
        For i = 1 To iSpaceNo
            If blnDirection Then
                StringFormat = StringFormat & " "
            Else
                StringFormat = " " & StringFormat
            End If
        Next i
    Else
        StringFormat = strSource
    End If
End Function

'strSource      要转换的字符串
'iLength        最终限制长度
'blnDirection   转换后空格位置 true:字符串前 false:字符串后
Private Function StringFormatSpace(strSource As String, iLength As Integer, bAlign As AlignEnum) As String
    Dim iStringLen As Integer
    Dim iSpaceNo As Integer
    Dim i As Integer
    Dim Schar(0 To 255) As String
    Dim HanZiCount As Integer
    Dim ZiFuCount As Integer

    iSpaceNo = iLength
    strSource = Trim(strSource)
    iStringLen = Len(strSource)
    For i = 0 To iStringLen - 1
        Schar(i) = Mid(strSource, i + 1, 1)
        If Asc(Schar(i)) < 0 Or Asc(Schar(i)) > 255 Then
            HanZiCount = HanZiCount + 1
            iSpaceNo = iSpaceNo - 2
        Else
            ZiFuCount = ZiFuCount + 1
            iSpaceNo = iSpaceNo - 1
        End If
        
        If iSpaceNo <= 0 Then
            Exit For
        End If
    Next i
    
    If iSpaceNo < 0 Then
        HanZiCount = HanZiCount - 1
        iSpaceNo = iSpaceNo + 2
    End If
    
    Select Case bAlign
        Case AlignEnum.emeLeft
            StringFormatSpace = Mid(strSource, 1, HanZiCount + ZiFuCount) + Space(iSpaceNo)
        Case AlignEnum.emeRight
            StringFormatSpace = Space(iSpaceNo) + Mid(strSource, 1, HanZiCount + ZiFuCount)
        Case AlignEnum.emeCenter
            StringFormatSpace = Space(iSpaceNo \ 2) + Mid(strSource, 1, HanZiCount + ZiFuCount) + Space(iSpaceNo - iSpaceNo \ 2)
        Case Else
    End Select
End Function

Private Function prnt11(X As Integer, y As Integer, Font As Single, Txt As String, Val As Integer)
    Dim str As String, str1 As String, str2 As String, i As Integer
    Dim distance As Integer
    
    Printer.CurrentX = X
    Printer.CurrentY = y
    Printer.FontBold = False
    Printer.FontSize = Font
    str = Txt
    str2 = str
    i = 0
    rowlab = 0
    
    Select Case Font
        Case 12
            distance = 240
        Case 13
            distance = 262
        Case 14
            distance = 281
        Case Else
    End Select
    
    If Len(Trim(str)) = 0 Then
        rowlab = 1   '待打印字符串为空的标志
    Else
        Do While Len(str) > 0
            Printer.CurrentX = X
            Printer.CurrentY = y + rowlab * distance
            rowlab = rowlab + 1
            If Len(str) >= Val Then
                str1 = Mid(str, 1, Val)
                Printer.Print str1
                i = i + 1
                str = Mid(str2, i * Val + 1)
            Else
                Printer.Print str
                Exit Do
            End If
        Loop
    End If
End Function
Private Function PrintFWAllByRow(vGrid As MSFlexGrid, ii As Integer) As String
    Dim strRowCol As String
    Dim strLine As String
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intMy As Integer, intStandar As Integer, intNet As Integer, intExeed As Integer
    Dim realMy As Integer, realExeed As Integer
    
        strLine = ""
        For j = 0 To conFixedCols - 1
        strRowCol = Trim(vGrid.TextMatrix(ii, j))
        Select Case j
            Case 0  '序号
                strRowCol = StringFormat(Trim(strRowCol), 4, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 4, AlignEnum.emeCenter)
            Case 1  '位号
                strRowCol = StringFormat(Trim(strRowCol), 4, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 4, AlignEnum.emeCenter)
            Case 2  '流量计值
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 3  '检尺值
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 4  '车号
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 8, AlignEnum.emeCenter)
            Case 5 '毛重
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 6 '皮重
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 7 '净重
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 8 '流-净
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 9 '表差/净
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 8, AlignEnum.emeCenter)
            Case 10 '尺差/净
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 8, AlignEnum.emeCenter)
        End Select
        Next j
    strLine = strLine + "│" + Space(12) + "│"
   PrintFWAllByRow = strLine
End Function
Private Function PrintFWAllByPage(vGrid As MSFlexGrid, ii As Integer) As String

End Function
Private Function PrintAllByPage(vGrid As MSFlexGrid, ii As Integer) As String
    Dim strRowCol As String
    Dim strLine As String
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intMy As Integer, intStandar As Integer, intNet As Integer, intExeed As Integer
    Dim realMy As Integer, realExeed As Integer
    
        strLine = ""
        For j = 0 To 4
        strRowCol = Trim(vGrid.TextMatrix(ii, j))
        Select Case j
            Case 0  '序号
                strRowCol = StringFormat(Trim(strRowCol), 4, True)
                strLine = strLine & "|" & StringFormatSpace(strRowCol, 4, AlignEnum.emeCenter)
            Case 1  '车型
            
            Case 2  '车号
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "|" & strRowCol
            Case 3  '总重
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "|" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 4  '速度
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "|" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
        End Select
        Next j
    strLine = strLine + "|" + Space(12) + "|"
   PrintAllByPage = strLine
End Function

Private Function PrintAllByRow(vGrid As MSFlexGrid, ii As Integer) As String
    Dim strRowCol As String
    Dim strLine As String
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intMy As Integer, intStandar As Integer, intNet As Integer, intExeed As Integer
    Dim realMy As Integer, realExeed As Integer
    
        strLine = ""
        For j = 0 To 4
        strRowCol = Trim(vGrid.TextMatrix(ii, j))
        Select Case j
            Case 0  '序号
                strRowCol = StringFormat(Trim(strRowCol), 4, True)
                strLine = strLine & "│" & StringFormatSpace(strRowCol, 4, AlignEnum.emeCenter)
            Case 1  '车型
            
            Case 2  '车号
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & strRowCol
            Case 3  '总重
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
            Case 4  '速度
                strRowCol = StringFormat(Trim(strRowCol), 8, True)
                strLine = strLine & "│" & StringFormatSpace(Format(strRowCol, "#0.00"), 8, AlignEnum.emeCenter)
        End Select
        Next j
    strLine = strLine + "│" + Space(12) + "│"
   PrintAllByRow = strLine
End Function

'行打印原始记录
Public Sub PrintOriginalDataByRow(vTime As String, vDirection As String, vGrid As MSFlexGrid)
    Dim i As Integer
    Dim intTotal As Single
    Dim strTotal As String
    Dim PrintString As String
     
    If vGrid.rows = 2 Or vGrid.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
    
    On Error GoTo Print_Err
    Open "lpt1" For Output As #3
    
    Printer.FontName = "宋体"
    Printer.FontSize = 18
    '调整行间距
    Print #3, Chr(27) & Chr(48)
    '调整列间距
    Print #3, Chr(28) & Chr$(83) & Chr$(0) & Chr$(0)
    
    Print #3, ""
    Print #3, ""
    Print #3, ""
    
    PrintString = StringFormatSpace(m_Corp + m_Tital, 50, AlignEnum.emeCenter)
    Print #3, PrintString
    Print #3, ""
    
    PrintString = " 日期: " + Mid(vTime, 1, 10) + Space(4) & "时间: " + Mid(vTime, 12, 8) + Space(4) + "操作员: " + g_LoginUser
    Print #3, PrintString
    
    PrintString = "┌──┬────┬────┬────┬──────┐"
    Print #3, PrintString
    
    PrintString = "│" + "序号" + "│" + " 车  号 " + "│" + " 重  量 " + "│" + " 速  度 " + "│" + "   备  注   " + "│"
    Print #3, PrintString
    
    PrintString = "│" + Space(4) + "│" + Space(8) + "│" + "  (t)  " + "│" + " (km/h) " + "│" + Space(12) + "│"
    Print #3, PrintString
    
    PrintString = "├──┼────┼────┼────┼──────┤"
    Print #3, PrintString
    
    For i = 1 To vGrid.rows - 2
        PrintString = PrintAllByRow(vGrid, i)
        Print #3, PrintString
        
        If i = vGrid.rows - 2 Then
            PrintString = "├──┴────┼────┼────┼──────┤"
            Print #3, PrintString
        Else
            PrintString = "├──┼────┼────┼────┼──────┤"
            Print #3, PrintString
        End If
    
        intTotal = intTotal + Val(Trim(vGrid.TextMatrix(i, 3)))    '总重累计
    Next i
    
    strTotal = str(intTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = "│" + StringFormatSpace("合计", 14, AlignEnum.emeCenter) + "│" + strTotal + "│" + Space(8) + "│" + Space(12) + "│"
    Print #3, PrintString
    
    PrintString = "└───────┴────┴────┴──────┘"
    Print #3, PrintString
    
    Print #3, ""
    Print #3, ""
    Print #3, ""
    
    Close #3
    
Print_Err:
    'MsgBox "打印错误，请检查打印机是否连接正确", vbOKOnly + vbInformation, "提示"
End Sub

'页打印原始记录
Public Sub PrintOriginalDataByPage(vTime As String, vDirection As String, vGrid As MSFlexGrid)
    Dim i As Integer, px As Integer, py As Integer
    Dim tt As Integer
    Dim intTotal As Single
    Dim strTotal As String
    Dim PrintString As String
    Dim strPrint As String
    
    If vGrid.rows = 2 Or vGrid.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
    
    px = 500
    py = 100
    Printer.FontName = "黑体"

    PrintString = StringFormatSpace(m_Corp + m_Tital, 40, AlignEnum.emeCenter)
    tt = prnt11(px, py, 12, PrintString, 110)
    PrintString = "日期: " + Mid(vTime, 1, 10) + Space(4) & "时间: " + Mid(vTime, 12, 8) + "操作员: " + g_LoginUser
    py = py + 300
    Printer.FontName = "宋体"
    tt = prnt11(px, py, 10, PrintString, 110)
    py = py + 250
    Printer.Line (px + 50, py)-(5955, py)
    py = py + 10
    PrintString = "|" + "序号" + "|" + " 车  号 " + "|" + " 重  量 " + "|" + " 速  度 " + "|" + "   备  注   " + "|"
    tt = prnt11(px, py, 12, PrintString, 110)
    PrintString = "|" + Space(4) + "|" + Space(8) + "|" + "  (t)  " + "|" + " (km/h) " + "|" + Space(12) + "|"
    py = py + 240
    tt = prnt11(px, py, 12, PrintString, 110)
    py = py + 250
    Printer.Line (px + 50, py)-(5955, py)
    For i = 1 To vGrid.rows - 2
        py = py + 10
        PrintString = PrintAllByPage(vGrid, i)
        tt = prnt11(px, py, 12, PrintString, 110)
        py = py + 240
        Printer.Line (px + 50, py)-(5955, py)
        intTotal = intTotal + Val(Trim(vGrid.TextMatrix(i, 3)))    '总重累计
    Next i
    strTotal = str(intTotal)
    strTotal = Format(strTotal, "#0.00")
    PrintString = "重量合计: " & strTotal & " 吨"
    py = py + 50
    tt = prnt11(px, py, 12, PrintString, 110)

    Printer.EndDoc
End Sub

Public Sub PrintOriginalData(vTime As String, vDirection As String, vGrid As MSFlexGrid)
    If m_Style = PrintStyle.emeRow Then
        PrintOriginalDataByRow vTime, vDirection, vGrid
    Else
        PrintOriginalDataByPage vTime, vDirection, vGrid
    End If
End Sub

'行打印对比数据
Public Sub PrintConstratDataByRow(vTime As String, vGrid As MSFlexGrid)
    Dim i As Integer
    Dim fmTotal, ruleTotal, grossTotal, tareTotal, netTotal, fnTotal As Single
    Dim strTotal As String
    Dim PrintString As String
    Dim header As String
     
    fmTotal = 0
    ruleTotal = 0
    grossTotal = 0
    tareTotal = 0
    netTotal = 0
    fnTotal = 0
    
    If vGrid.rows = 2 Or vGrid.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
    
    On Error GoTo Print_Err
    Open "lpt1" For Output As #3
    
    Printer.FontName = "宋体"
    Printer.FontSize = 18
    '调整行间距
    Print #3, Chr(27) & Chr(48)
    '调整列间距
    Print #3, Chr(28) & Chr$(83) & Chr$(0) & Chr$(0)
    
    Print #3, ""
    Print #3, ""
    Print #3, ""
    
    header = m_Corp + "轨道衡静态计量与流量计比对报表"
    PrintString = StringFormatSpace(header, 100, AlignEnum.emeCenter)
    Print #3, PrintString
    Print #3, ""
    
    PrintString = " 日期: " + Mid(vTime, 1, 10) + Space(4) & "时间: " + Mid(vTime, 12, 8) + "操作员: " + g_LoginUser
    Print #3, PrintString
    
    PrintString = "┌──┬──┬────┬────┬────┬────┬────┬────┬────┬────┬────┐"
    Print #3, PrintString
    
    PrintString = "│" + "序号" + "│" + "位号" + "│" + " 流  量 " + "│" + " 检  尺 " + "│" + " 车  号 " + "│" + " 毛  重 " + "│" + " 皮  重 " + "│" + " 静  重 " + "│" + " 流-净  " + "│" + " 表差/  " + "│" + " 尺差/  " + "│"
    Print #3, PrintString
    
    PrintString = "│" + Space(4) + "│" + Space(4) + "│" + " 计  值 " + "│" + "   值   " + "│" + Space(8) + "│" + "  (t)   " + "│" + "  (t)   " + "│" + "  (t)   " + "│" + "  (t)   " + "│" + " 净(‰) " + "│" + " 净(‰) " + "│"
    Print #3, PrintString
    
    PrintString = "├──┼──┼────┼────┼────┼────┼────┼────┼────┼────┼────┤"
    Print #3, PrintString
    
    For i = 1 To vGrid.rows - 2
        PrintString = PrintFWAllByRow(vGrid, i)
        Print #3, PrintString
        
        If i <> vGrid.rows - 2 Then
            PrintString = "├──┼──┼────┼────┼────┼────┼────┼────┼────┼────┼────┤"
        Else
            PrintString = "├──┴──┼────┼────┼────┼────┼────┼────┼────┼────┼────┤"
        End If
        Print #3, PrintString
    
        '总重累计
        fmTotal = fmTotal + Val(Trim(vGrid.TextMatrix(i, 2)))
        ruleTotal = ruleTotal + Val(Trim(vGrid.TextMatrix(i, 3)))
        grossTotal = grossTotal + Val(Trim(vGrid.TextMatrix(i, 5)))
        tareTotal = tareTotal + Val(Trim(vGrid.TextMatrix(i, 6)))
        netTotal = netTotal + Val(Trim(vGrid.TextMatrix(i, 7)))
        fnTotal = fnTotal + Val(Trim(vGrid.TextMatrix(i, 8)))
    Next i
    
    strTotal = str(fmTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = "│" + StringFormatSpace("合计", 10, AlignEnum.emeCenter) + "│" + strTotal + "│"
    
    strTotal = str(ruleTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    PrintString = PrintString + Space(8) + "│"
    
    strTotal = str(grossTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(tareTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(netTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(fnTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    PrintString = PrintString + Space(8) + "│" + Space(8) + "│"
    Print #3, PrintString
    
    PrintString = "└─────┴────┴────┴────┴────┴────┴────┴────┴────┴────┘"
    Print #3, PrintString
    
    Print #3, ""
    Print #3, ""
    Print #3, ""
    
    Close #3
    
Print_Err:
    'MsgBox "打印错误，请检查打印机是否连接正确", vbOKOnly + vbInformation, "提示"
End Sub
' 页打印对比数据
Public Sub PrintConstratDataByPage(vTime As String, vGrid As MSFlexGrid)
    Dim i As Integer, px As Integer, py As Integer
    Dim tt As Integer
    Dim intTotal As Single
    Dim strTotal As String
    Dim PrintString As String
    Dim strPrint As String
    Dim header As String
    
    If vGrid.rows = 2 Or vGrid.TextMatrix(1, 0) = "" Then
        Exit Sub
    End If
    
    px = 500
    py = 100
    Printer.FontName = "黑体"
    
    header = m_Corp + m_Tital
    PrintString = StringFormatSpace(header, 40, AlignEnum.emeCenter)
    tt = prnt11(px, py, 12, PrintString, 110)
    PrintString = "日期: " + Mid(vTime, 1, 10) + Space(4) & "时间: " + Mid(vTime, 12, 8) + "操作员: " + g_LoginUser
    py = py + 300
    Printer.FontName = "宋体"
    tt = prnt11(px, py, 10, PrintString, 110)
    py = py + 250
    Printer.Line (px + 50, py)-(5955, py)
    py = py + 10
    
    PrintString = "│" + "序号" + "│" + "位号" + "│" + " 流  量 " + "│" + " 检  尺 " + "│" + " 车  号 " + "│" + " 毛  重 " + "│" + " 皮  重 " + "│" + " 静  重 " + "│" + " 流-净  " + "│" + " 表差/  " + "│" + " 尺差/  " + "│"
    tt = prnt11(px, py, 12, PrintString, 110)
    PrintString = "│" + Space(4) + "│" + Space(4) + "│" + " 计  值 " + "│" + "   值   " + "│" + Space(8) + "│" + "  (t)   " + "│" + "  (t)   " + "│" + "  (t)   " + "│" + "  (t)   " + "│" + " 净(‰) " + "│" + " 净(‰) " + "│"
    py = py + 240
    tt = prnt11(px, py, 12, PrintString, 110)
    py = py + 250
    Printer.Line (px + 50, py)-(5955, py)
    For i = 1 To vGrid.rows - 2
        py = py + 10
        PrintString = PrintFWAllByRow(vGrid, i)
        tt = prnt11(px, py, 12, PrintString, 110)
        py = py + 240
        Printer.Line (px + 50, py)-(5955, py)
        
        '总重累计
        fmTotal = fmTotal + Val(Trim(vGrid.TextMatrix(i, 2)))
        ruleTotal = ruleTotal + Val(Trim(vGrid.TextMatrix(i, 3)))
        grossTotal = grossTotal + Val(Trim(vGrid.TextMatrix(i, 5)))
        tareTotal = tareTotal + Val(Trim(vGrid.TextMatrix(i, 6)))
        netTotal = netTotal + Val(Trim(vGrid.TextMatrix(i, 7)))
        fnTotal = fnTotal + Val(Trim(vGrid.TextMatrix(i, 8)))
    Next i
    
    strTotal = str(fmTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = "│" + StringFormatSpace("合计", 10, AlignEnum.emeCenter) + "│" + strTotal + "│"
    
    strTotal = str(ruleTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    PrintString = PrintString + Space(8) + "│"
    
    strTotal = str(grossTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(tareTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(netTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    strTotal = str(fnTotal)
    strTotal = StringFormatSpace(Format(strTotal, "#0.00"), 8, AlignEnum.emeCenter)
    PrintString = PrintString + strTotal + "│"
    
    PrintString = PrintString + Space(8) + "│" + Space(8) + "│"

    py = py + 50
    tt = prnt11(px, py, 12, PrintString, 110)
    
    py = py + 240
    Printer.Line (px + 50, py)-(5955, py)

    Printer.EndDoc
End Sub
Public Sub PrintConstratData(vTime As String, vGrid As MSFlexGrid)
    If m_Style = PrintStyle.emeRow Then
        PrintConstratDataByRow vTime, vGrid
    Else
        PrintConstratDataByPage vTime, vGrid
    End If
End Sub

'打印检衡记录
'Public Sub PrintDetectData(vTime As String, vDirection As String, vGrid As MSFlexGrid)
'    Dim strP As String
'    Dim I As Integer
'
'    On Error GoTo Print_Err
'    Open "lpt1" For Output As #3
'        Print #3, "             检  衡  单"
'        Print #3, "========" & vTime & "==========="
'
'        For I = 1 To vGrid.Rows - 2
'            strP = StringFormat(Trim(vGrid.TextMatrix(I, 0)), 8, True)
'            strP = strP & StringFormat(Trim(vGrid.TextMatrix(I, 3)), 12, True)
'            strP = strP & StringFormat(Trim(vGrid.TextMatrix(I, 4)), 12, True)
'
'            Print #3, strP
'            strP = ""
'        Next I
'
'        Print #3, "========" & vTime & "==========="
'        Print #3, ""
'    Close #3
'
'Print_Err:
'End Sub

'打印检衡记录
Public Sub PrintDetectData(vTime As String, vDirection As String, vGrid As MSFlexGrid)
    Dim strP As String
    Dim i As Integer
    Dim totalWeight As Single
    
    On Error GoTo Print_Err
    Open "lpt1" For Output As #3
'    Open App.Path & "\JHdata.TXT" For Append As #3
        Print #3, "                     检  衡  单"
        Print #3, "================" & vTime & "==================="

        For i = 1 To vGrid.rows - 2
            strP = StringFormat(Trim(vGrid.TextMatrix(i, 0)), 8, True)
            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 1)), 8, True)
            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 2)), 14, True)
            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 3)), 12, True)
            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 4)), 12, True)
            totalWeight = totalWeight + Val(vGrid.TextMatrix(i, 3))
            Print #3, strP
            strP = ""
        Next i

        Print #3, "================" & vTime & "==================="
        Print #3, "总重：" & totalWeight
        Print #3, ""
    Close #3
    
    Open App.Path & "\BRW.OBL" For Append As #31

    
        For i = 1 To vGrid.rows - 2
           TotalXH = TotalXH + 1
'            strP = StringFormat(Trim(vGrid.TextMatrix(i, 0)), 8, True)
'            strP = StringFormat(Trim(vGrid.TextMatrix(i, 1)), 8, True)
'            strP = StringFormat(Trim(vGrid.TextMatrix(i, 2)), 8, True)
'            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 3)), 12, True)
'            strP = strP & StringFormat(Trim(vGrid.TextMatrix(i, 4)), 12, True)
'            TotalWeight = TotalWeight + Val(vGrid.TextMatrix(i, 3))
            strP = str(TotalXH) & Chr(9) & Trim(vGrid.TextMatrix(i, 2)) & Chr(9) & Val(vGrid.TextMatrix(i, 3)) * 1000 & Chr(9) & Trim(vGrid.TextMatrix(i, 4)) & Chr(9) & vTime
            Debug.Print strP
            Print #31, strP
            strP = ""
        Next i

   
'        Print #31, ""
    Close #31
Print_Err:
End Sub

'打印实验报告
Public Sub PrintReportData(vLeftCarriage As CDebugCarriage, vRightCarriage As CDebugCarriage, vStander() As Single)
    Dim PrintString As String
    Dim i As Integer, px As Integer, py As Integer
    Dim StanderNum As Integer
    Dim StanderWeight(1 To 5) As Single
    Dim LDiffWeight(1 To 5) As Single
    Dim RDiffWeight(1 To 5) As Single
    Dim LScaleWeight(1 To 5) As Single
    Dim RScaleWeight(1 To 5) As Single
    Dim EndIndex As Integer
    Dim FileNo As Integer
    Dim FilePath As String
    
    FilePath = App.Path + "\检测报告.txt"
    
    If Dir(FilePath) <> "" Then
        Kill FilePath
    End If
    
    FileNo = FreeFile
    Open FilePath For Output As #FileNo
    
    'vStander数组下标是从0开始的,但是vStander[0]未使用
    On Error GoTo UBound_err
    StanderNum = UBound(vStander) - LBound(vStander)
UBound_err:
    If Err.Number <> 0 Then StanderNum = 0
    
    If StanderNum <= 5 And StanderNum > 0 Then
        EndIndex = UBound(vStander)
    ElseIf StanderNum <= 0 Then
        EndIndex = 0
    Else
        EndIndex = 5
    End If
    
    For i = 1 To EndIndex
        StanderWeight(i) = vStander(i)
    Next i
    
    
    Printer.Orientation = 2
    Printer.FontName = "黑体"

    px = 0
    py = 100
    '标题
    PrintString = Space(30) & "动态电子轨道衡重复性实验报告(" & Format(Now, "yyyy年mm月dd日") & ")"
    tt = prnt11(px, py, 16, PrintString, 110)
    Print #FileNo, PrintString
    
    Printer.FontName = "宋体"
    py = py + 300
    PrintString = strLeftTop + StringConnection(strHLine, 8) + strCenterTop + StringConnection(strHLine, 24) + strCenterTop + StringConnection(strHLine, 3) + strCenterTop + StringConnection(strHLine, 24) + strRightTop
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("方向", 16, emeCenter) + strVLine + StringFormatSpace("左", 48, emeCenter) + strVLine + StringFormatSpace("", 6, emeCenter) + strVLine + StringFormatSpace("右", 48, emeCenter) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strLeftCenter + StringConnection(strHLine, 8) + strCenter + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("车号", 16, emeCenter) + strVLine + StringFormatSpace(vLeftCarriage.Code(1), 8, emeCenter) + strVLine + StringFormatSpace(vLeftCarriage.Code(2), 8, emeCenter) + strVLine + StringFormatSpace(vLeftCarriage.Code(3), 8, emeCenter) + strVLine + StringFormatSpace(vLeftCarriage.Code(4), 8, emeCenter) + strVLine + StringFormatSpace(vLeftCarriage.Code(5), 8, emeCenter) + strVLine + StringFormatSpace("", 6, emeCenter) + strVLine + StringFormatSpace(vRightCarriage.Code(1), 8, emeCenter) + strVLine + StringFormatSpace(vRightCarriage.Code(2), 8, emeCenter) + strVLine + StringFormatSpace(vRightCarriage.Code(3), 8, emeCenter) + strVLine + StringFormatSpace(vRightCarriage.Code(4), 8, emeCenter) + strVLine + StringFormatSpace(vRightCarriage.Code(5), 8, emeCenter) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    '两行之间的间隔
    py = py + 240
    PrintString = strLeftCenter + StringConnection(strHLine, 4) + strCenterTop + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    For i = 1 To 10
        py = py + 240
        PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace(Format(i, "#00"), 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.Weight(i, 1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.Weight(i, 2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.Weight(i, 3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.Weight(i, 4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.Weight(i, 5), "#0.000")), 8, emeRight) + strVLine _
                    + StringFormatSpace(Format(i, "#00"), 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.Weight(i, 1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.Weight(i, 2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.Weight(i, 3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.Weight(i, 4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.Weight(i, 5), "#0.000")), 8, emeRight) + strVLine
        tt = prnt11(px, py, 12, PrintString, 1100)
        Print #FileNo, PrintString
        py = py + 240
        If i = 10 Then
            PrintString = strVLine + StringFormatSpace("示值", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
        Else
            PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
        End If
        tt = prnt11(px, py, 12, PrintString, 1100)
        Print #FileNo, PrintString
    Next i
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("最大值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MaxWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MaxWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MaxWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MaxWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MaxWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("最大值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MaxWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MaxWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MaxWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MaxWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MaxWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    If vLeftCarriage.Count > 0 Then
        For i = 1 To EndIndex
            LDiffWeight(i) = vLeftCarriage.MaxWeight(i) - StanderWeight(i)
        Next i
    End If
    If vRightCarriage.Count > 0 Then
        For i = 1 To EndIndex
            RDiffWeight(i) = vRightCarriage.MaxWeight(i) - StanderWeight(EndIndex - i + 1)
        Next i
    End If
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    For i = 1 To EndIndex
        If Abs(StanderWeight(i)) >= 0.000001 Then
            LScaleWeight(i) = vLeftCarriage.MaxWeight(i) / StanderWeight(i)
            RScaleWeight(i) = vRightCarriage.MaxWeight(i) / StanderWeight(EndIndex - i + 1)
        End If
    Next i
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("最小值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MinWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MinWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MinWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MinWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.MinWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("最小值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MinWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MinWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MinWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MinWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.MinWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    If vLeftCarriage.Count > 0 Then
        For i = 1 To EndIndex
            LDiffWeight(i) = vLeftCarriage.MinWeight(i) - StanderWeight(i)
        Next i
    End If
    If vRightCarriage.Count > 0 Then
        For i = 1 To EndIndex
            RDiffWeight(i) = vRightCarriage.MinWeight(i) - StanderWeight(EndIndex - i + 1)
        Next i
    End If
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    For i = 1 To EndIndex
        If Abs(StanderWeight(i)) >= 0.000001 Then
            LScaleWeight(i) = vLeftCarriage.MinWeight(i) / StanderWeight(i)
            RScaleWeight(i) = vRightCarriage.MinWeight(i) / StanderWeight(EndIndex - i + 1)
        End If
    Next i
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("平均值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.AvgWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.AvgWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.AvgWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.AvgWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vLeftCarriage.AvgWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("平均值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.AvgWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.AvgWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.AvgWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.AvgWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(vRightCarriage.AvgWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    If vLeftCarriage.Count > 0 Then
        For i = 1 To EndIndex
            LDiffWeight(i) = vLeftCarriage.AvgWeight(i) - StanderWeight(i)
        Next i
    End If
    If vRightCarriage.Count > 0 Then
        For i = 1 To EndIndex
            RDiffWeight(i) = vRightCarriage.AvgWeight(i) - StanderWeight(EndIndex - i + 1)
        Next i
    End If
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LDiffWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("差值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RDiffWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    For i = 1 To EndIndex
        If Abs(StanderWeight(i)) >= 0.000001 Then
            LScaleWeight(i) = vLeftCarriage.AvgWeight(i) / StanderWeight(i)
            RScaleWeight(i) = vRightCarriage.AvgWeight(i) / StanderWeight(EndIndex - i + 1)
        End If
    Next i
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(LScaleWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("比值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(RScaleWeight(5), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strLeftCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 3) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strCenter + StringConnection(strHLine, 4) + strRightCenter
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 240
    PrintString = strVLine + StringFormatSpace("", 8, emeCenter) + strVLine + StringFormatSpace("标准值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(1), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(5), "#0.000")), 8, emeRight) + strVLine _
                + StringFormatSpace("标准值", 6, emeCenter) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(5), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(4), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(3), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(2), "#0.000")), 8, emeRight) + strVLine + StringFormatSpace(StringChange(Format(StanderWeight(1), "#0.000")), 8, emeRight) + strVLine
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    py = py + 240
    PrintString = strLeftButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 3) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 3) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strCenterButtom + StringConnection(strHLine, 4) + strRightButtom
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    py = py + 300
    PrintString = " 说明：差值是最大值、最小值或平均值与标准值的差。       比例是最大值、最小值或平均值与标准值的比"
    tt = prnt11(px, py, 12, PrintString, 1100)
    Print #FileNo, PrintString
    
    Close #FileNo
    Printer.EndDoc
End Sub

Private Sub UserControl_Initialize()
    Dim mConfig As CConfig
    
    Set mConfig = New CConfig
    
    With mConfig
        .FileName = App.Path + "\gdh.bin"
    End With
    
    m_Style = mConfig.GetInteger("print", "PrintStyle", PrintStyle.emePage)
    m_Tital = mConfig.GetString("print", "tital", "静态电子轨道衡称重计量单")
    m_Corp = mConfig.GetString("print", "corp", "北京东方瑞威")
End Sub
