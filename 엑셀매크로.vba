Sub ��ũ��1()
    '
' ��ũ��1 ��ũ��
'
' �ٷ� ���� Ű: Ctrl+q
'
Dim temp
Dim temp2

    Range("A1:H39").Select
    Selection.ClearContents
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "�޴ºм���"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "�޴º���ȭ��ȣ"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "�޴ºб�Ÿ����ó"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "�޴ºп����ȣ"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "�޴º��ּ�(��ü, ����)"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "ǰ���"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "��ǰ����"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "��۸޼���1"

Dim excel(4) As String
excel(0) = "��۸���.xls"
excel(1) = Dir("C:\Users\Administrator\Desktop\����Ʈ�����*")
excel(2) = Dir("C:\Users\Administrator\Desktop\*logistics*")
excel(3) = Dir("C:\Users\Administrator\Desktop\GeneralDelivery*")
excel(4) = Dir("C:\Users\Administrator\Desktop\�߼۰���*")

Dim arr(4) As Variant
arr(0) = Array("�޴ºм���", "�޴º���ȭ��ȣ", "�޴ºб�Ÿ����ó", "�޴ºп����ȣ", "�޴º��ּ�(��ü, ����)", "ǰ���", "��ǰ����", "��۸޼���1", "�ɼ�")
arr(1) = Array("�����θ�", "�����ο���ó1", "�����ο���ó2", "�����ȣ", "�����", "��ǰ��", "����", "��۸޼���", "�ɼ�����")
arr(2) = Array("������", "�޴�����ȣ", "��ȭ��ȣ", "�����ȣ", "�ּ�", "��ǰ��", "����", "��۸޽���", "�ɼ�")
arr(3) = Array("�����θ�", "������ �޴���", "������ ��ȭ��ȣ", "�����ȣ", "�ּ�", "��ǰ��", "����", "��۽� �䱸����", "�ֹ��ɼ�")
arr(4) = Array("������", "��������ȭ��ȣ", "�������ڵ�����ȣ", "�����ο����ȣ", "�������ּ�", "��ǰ��", "����", "��۸޽���", "��ǰ�ɼ�")





    Range("I1").Select
    ActiveCell.FormulaR1C1 = "�ɼ�"
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
    Cells.Find(What:="�Ǹ��ڻ�ǰ�ɼǹ�ȣ", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
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
    
    Cells.Find(What:="ǰ���", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
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
