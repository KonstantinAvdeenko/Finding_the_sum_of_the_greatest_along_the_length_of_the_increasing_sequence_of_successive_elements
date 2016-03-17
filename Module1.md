Private Sub CommandButton1_Click()
'заполняет элементы массива рандомными положительными значениями'
For i = 1 To 30
Cells(1, i) = Int((30 * Rnd) + 1)
Next i
End Sub

Private Sub CommandButton2_Click()
'находит и выводит сумму наибольшей по длине возрастающей последовательности подряд идущих элементов'
Max = 0
imax = 0
For i = 1 To 29
If Cells(1, i) < Cells(1, i + 1) Then
k = k + 1
Else
k = k + 1 'Один элемент последовательности будет всегда'
    If Max < k Then
        Max = k
        imax = i
    End If
k = 0 'Обнуление длины возврастающей последовательности по окончанию возврастания'
End If
Next i
MsgBox (Max) 'Показ длины возврастающей последовательности'
MsgBox (imax) 'Показ количества возрастающих последовательностей'
End Sub

Private Sub CommandButton4_Click()
'Закрытие формы'
UserForm1.Hide
End Sub