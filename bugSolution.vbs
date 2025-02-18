Function f(a, b)
  'Explicitly convert both a and b to numbers before comparison
  Dim aNum, bNum
  aNum = CDbl(a) 
  bNum = CDbl(b)

  If aNum > bNum Then
    MsgBox "a is greater than b"
  ElseIf aNum < bNum Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end function