Attribute VB_Name = "ģ��2"
Option Explicit

Sub ���ڵļ���()
Attribute ���ڵļ���.VB_Description = "���ڼ�����ڵ�ֵ"
Attribute ���ڵļ���.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' ���ڵļ��� ��
' ���ڼ�����ڵ�ֵ
'
' ��ݼ�: Ctrl+Shift+F
'
    '����ʱ������
    Dim time As Double, cash As Double, x As Integer, rate As Double, pv As Double, number As Integer, danyuange As String
    'ѡ����ĳ����Ԫ��Ŀ�ʼ�����򣬲���������ݺ͸�ʽ
    range(Cells(5, 1), Cells.SpecialCells(xlCellTypeLastCell)).Clear
    Cells(1, 2).Value = InputBox("������ʱ������")
    Cells(1, 4).Value = InputBox("������ÿ�ڵ��ֽ���")
    Cells(2, 2).Value = InputBox("������ÿ�ڵ�����|�����ʣ�С����ʾ")
    Cells(2, 4).Value = InputBox("��������ֵ")
    time = Cells(1, 2).Value
    cash = Cells(1, 4).Value
    rate = Cells(2, 2).Value
    number = Cells(2, 4).Value
    ' Debug.Print cash
     
    For x = 1 To time
        If x >= time Then
        Cells(4 + x, 1).Value = x
        Cells(4 + x, 2).Value = cash + number
        pv = (Cells(4 + x, 2).Value) / (1 + rate) ^ x
        Cells(4 + x, 3).Value = pv
        
        Cells(5 + x, 1).Value = "�ϼ�"
        range(Cells(5 + x, 1), Cells(5 + x, 2)).Merge
        
        
        
        Cells(5 + x, 3) = Application.WorksheetFunction.Sum(range(Cells(5, 3), Cells(4 + time, 3)))
        Else
            Cells(4 + x, 1).Value = x
            Cells(4 + x, 2).Value = cash
            pv = cash / (1 + rate) ^ x
            Cells(4 + x, 3).Value = pv
            
        End If
    Next x
    For x = 1 To time
        If x >= time Then
        Cells(4 + x, 4) = Cells(4 + x, 3) / Cells(5 + time, 3)
        Cells(4 + x, 5) = Cells(4 + x, 1) * Cells(4 + x, 4)
        Cells(5 + x, 4) = Application.WorksheetFunction.Sum(range(Cells(5, 4), Cells(4 + time, 4)))
        Cells(5 + x, 5) = Application.WorksheetFunction.Sum(range(Cells(5, 5), Cells(4 + time, 5)))
        Cells(3, 2) = Cells(5 + x, 5)
        
        Else
        Cells(4 + x, 4) = Cells(4 + x, 3) / Cells(5 + time, 3)
        Cells(4 + x, 5) = Cells(4 + x, 1) * Cells(4 + x, 4)
        End If
        
    Next x
    
    
    
End Sub
