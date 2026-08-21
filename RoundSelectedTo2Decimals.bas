Attribute VB_Name = "RoundSelectedTo2Decimals"
Sub RoundSelectedTo2Decimals()
    Dim cell As Range
    Dim numVal As Double
    Dim txt As String
    Dim processed As String

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False

    For Each cell In Selection.Cells
        If IsEmpty(cell) Then GoTo NextCell
        If IsError(cell.Value) Then GoTo NextCell
        If IsDate(cell.Value) Then GoTo NextCell

        ' 默认跳过公式单元格；如果不想跳过，可删除下一行
        If cell.HasFormula Then GoTo NextCell

        ' 数值单元格，包括文本格式的纯数字
        If IsNumeric(cell.Value) Then
            numVal = CDbl(cell.Value)

            If InStr(1, cell.NumberFormat, "%", vbTextCompare) > 0 Then
                ' 百分比格式：直接对百分号前的数字保留两位
                ' 因此值本身保留 4 位小数
                cell.Value = Application.WorksheetFunction.Round(numVal, 4)
                cell.NumberFormat = "0.00%"
            Else
                cell.Value = Application.WorksheetFunction.Round(numVal, 2)
                cell.NumberFormat = "0.00"
            End If
        ElseIf VarType(cell.Value) = vbString Then
            txt = cell.Text
            processed = RoundFirstNumberInText(txt)
            cell.Value = processed
        End If

NextCell:
    Next cell

    Application.ScreenUpdating = True
    MsgBox "处理完成。", vbInformation
End Sub

Function RoundFirstNumberInText(ByVal txt As String) As String
    Dim reg As Object
    Dim matches As Object
    Dim m As Object
    Dim numText As String
    Dim cleanNum As String
    Dim val As Double
    Dim roundedText As String

    Set reg = CreateObject("VBScript.RegExp")

    ' 匹配数字：可选负号、千分位逗号、小数点，也支持 .5 这类写法
    reg.Pattern = "-?(?:\d[\d,]*(?:\.\d+)?|\.\d+)"
    reg.Global = False

    Set matches = reg.Execute(txt)
    If matches.Count = 0 Then
        RoundFirstNumberInText = txt
        Exit Function
    End If

    Set m = matches(0)
    numText = m.Value
    cleanNum = Replace(numText, ",", "")

    If Not IsNumeric(cleanNum) Then
        RoundFirstNumberInText = txt
        Exit Function
    End If

    val = Application.WorksheetFunction.Round(CDbl(cleanNum), 2)

    ' 如果原数字带千分位逗号，结果也保留千分位
    If InStr(numText, ",") > 0 Then
        roundedText = Format(val, "#,##0.00")
    Else
        roundedText = Format(val, "0.00")
    End If

    ' 只替换第一个数字，后面的文字、百分号等原样保留
    RoundFirstNumberInText = Left(txt, m.FirstIndex) & roundedText & Mid(txt, m.FirstIndex + m.Length + 1)
End Function