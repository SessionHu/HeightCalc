h = InputBox("请输入身高，单位随意")
If h < 10 Then
 r = h*100
 MsgBox r ,, "身高（cm）"
Else
 r = h/100
 MsgBox r ,, "身高（m）"
End If