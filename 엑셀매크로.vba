Sub 매크로1()
    '
' 매크로1 매크로
'
' 바로 가기 키: Ctrl+q
'
Dim temp
Dim temp2

    Range("A1:H39").Select
    Selection.ClearContents
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "받는분성명"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "받는분전화번호"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "받는분기타연락처"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "받는분우편번호"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "받는분주소(전체, 분할)"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "품목명"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "내품수량"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "배송메세지1"

Dim excel(4) As String
excel(0) = "배송명세서.xls"
excel(1) = Dir("C:\Users\Administrator\Desktop\스마트스토어*")
excel(2) = Dir("C:\Users\Administrator\Desktop\*logistics*")
excel(3) = Dir("C:\Users\Administrator\Desktop\GeneralDelivery*")
excel(4) = Dir("C:\Users\Administrator\Desktop\발송관리*")

Dim arr(4) As Variant
arr(0) = Array("받는분성명", "받는분전화번호", "받는분기타연락처", "받는분우편번호", "받는분주소(전체, 분할)", "품목명", "내품수량", "배송메세지1", "옵션")
arr(1) = Array("수취인명", "수취인연락처1", "수취인연락처2", "우편번호", "배송지", "상품명", "수량", "배송메세지", "옵션정보")
arr(2) = Array("수취인", "휴대폰번호", "전화번호", "우편번호", "주소", "상품명", "수량", "배송메시지", "옵션")
arr(3) = Array("수령인명", "수령인 휴대폰", "수령인 전화번호", "우편번호", "주소", "상품명", "수량", "배송시 요구사항", "주문옵션")
arr(4) = Array("수령인", "수령인전화번호", "수령인핸드폰번호", "수령인우편번호", "수령인주소", "상품명", "수량", "배송메시지", "상품옵션")





    Range("I1").Select
    ActiveCell.FormulaR1C1 = "옵션"
If excel(1) = "" Then
Else
    Workbooks.Open ("C:\Users\Administrator\Desktop\" + excel(1))
    For i = 0 To 8
        Windows(excel(1)).Activate
        
        Cells.Find(What:=arr(1)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
        
        
        ActiveCell.Offset(1, 0).Range("A1:A50").Select
        Selection.Copy
        
        Windows(excel(0)).Activate
        
        Cells.Find(What:=arr(0)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
    Next i
End If

If excel(2) = "" Then
Else
    Workbooks.Open ("C:\Users\Administrator\Desktop\" + excel(2))
     For i = 0 To 8
        Windows(excel(2)).Activate
        
        Cells.Find(What:=arr(2)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
        
        
        ActiveCell.Offset(1, 0).Range("A1:A50").Select
        Selection.Copy
        
        Windows(excel(0)).Activate
        
        Cells.Find(What:=arr(0)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(1, 0).Range("A51").Select
        ActiveSheet.Paste
    Next i
End If
    
If excel(3) = "" Then
Else
    Workbooks.Open ("C:\Users\Administrator\Desktop\" + excel(3))
     For i = 0 To 8
        Windows(excel(3)).Activate
        
        Cells.Find(What:=arr(3)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
        
        
        ActiveCell.Offset(1, 0).Range("A1:A50").Select
        Selection.Copy
        
        Windows(excel(0)).Activate
        
        Cells.Find(What:=arr(0)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(1, 0).Range("A101").Select
        ActiveSheet.Paste
    Next i
End If
If excel(4) = "" Then
Else
    Workbooks.Open ("C:\Users\Administrator\Desktop\" + excel(4))
    Cells.Find(What:="판매자상품옵션번호", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, MatchByte:=False, SearchFormat:=False).Activate
    Selection.ClearContents
     For i = 0 To 8
        Windows(excel(4)).Activate
        
        Cells.Find(What:=arr(4)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
        
        
        ActiveCell.Offset(1, 0).Range("A1:A50").Select
        Selection.Copy
        
        Windows(excel(0)).Activate
        
        Cells.Find(What:=arr(0)(i), After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , MatchByte:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(1, 0).Range("A151").Select
        ActiveSheet.Paste
    Next i
End If
    
    Cells.Find(What:="품목명", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, MatchByte:=False, SearchFormat:=False).Activate
    
    For i = 0 To 250

        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveCell.Offset(0, 3).Range("A1").Select
        temp = Selection.Value
        ActiveCell.Offset(0, -3).Range("A1").Select
        temp2 = Selection.Value
        ActiveCell.FormulaR1C1 = temp2 + " " + temp
    Next i
    
    Range("I1:I200").Select
    Selection.ClearContents
    
    Range("A250").Select

    For i = 2 To 250
    
    If Selection.Value = "" Then
        ActiveCell.Offset(0, 0).Rows("1:1").EntireRow.Select
        Selection.Delete Shift:=xlUp
    End If
    ActiveCell.Offset(-1, 0).Range("A1").Select
Next i
End Sub
