h = InputBox("��������ߣ���λ����")
If h < 10 Then
 r = h*100
 MsgBox r ,, "��ߣ�cm��"
Else
 r = h/100
 MsgBox r ,, "��ߣ�m��"
End If