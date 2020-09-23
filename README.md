<div align="center">

## Evaluate Equation


</div>

### Description

Evaluate any Mathematical Equation

like: Evaluate("(5+3)*2")
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lefteris Eleftheriades](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lefteris-eleftheriades.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lefteris-eleftheriades-evaluate-equation__1-61479/archive/master.zip)





### Source Code

```
Function Evaluate(ByVal Equation As String)
 Dim i&, P%, Si&
 Equation = Replace(Replace(Equation, ",", "."), "--", "")
 'Parenthesis Begins with outermost
 For i = 1 To Len(Equation)
  If Mid(Equation, i, 1) = "(" Then
  P = P + 1
  If P = 1 Then Si = i
  End If
  If Mid(Equation, i, 1) = ")" Then
  If P = 1 Then
   Evaluate = Evaluate(Mid(Equation, 1, Si - 1) & Evaluate(Mid(Equation, Si + 1, i - 1 - Si)) & Mid(Equation, i + 1))
   Exit Function
  End If
  P = P - 1
  End If
 Next
 'Addition / Substruction
  For i = Len(Equation) To 1 Step -1
   If Mid(Equation, i, 1) = "+" Then
    If i > 1 Then
      If Mid(Equation, i - 1, 1) <> "*" And Mid(Equation, i - 1, 1) <> "/" Then
       Evaluate = Evaluate(Mid(Equation, 1, i - 1)) + Evaluate(Mid(Equation, i + 1))
       Exit Function
      End If
    Else
      Evaluate = Evaluate(Mid(Equation, 1, i - 1)) + Evaluate(Mid(Equation, i + 1))
      Exit Function
    End If
   End If
   If Mid(Equation, i, 1) = "-" Then
    If i > 1 Then
      If Mid(Equation, i - 1, 1) <> "*" And Mid(Equation, i - 1, 1) <> "/" Then
       Evaluate = Evaluate(Mid(Equation, 1, i - 1)) - Evaluate(Mid(Equation, i + 1))
       Exit Function
      End If
    Else
      Evaluate = Evaluate(Mid(Equation, 1, i - 1)) - Evaluate(Mid(Equation, i + 1))
      Exit Function
    End If
   End If
  Next
 'Multiplication / Division
 For i = Len(Equation) To 1 Step -1
  If Mid(Equation, i, 1) = "*" Then
  Evaluate = Evaluate(Mid(Equation, 1, i - 1)) * Evaluate(Mid(Equation, i + 1))
  Exit Function
  End If
  If Mid(Equation, i, 1) = "/" Then
  Evaluate = Evaluate(Mid(Equation, 1, i - 1)) / Evaluate(Mid(Equation, i + 1))
  Exit Function
  End If
 Next
 Evaluate = Val(Equation)
End Function
```

