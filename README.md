<div align="center">

## Base Conversion Pro


</div>

### Description

The function BaseConv convert any base 10 numbers to any selected base (max 64) and the function ConvBase10 convert it back to base 10.

It can be useful for cryptography purpose example: convert the ASCII value of a char to a secret base...or it can be useful for any algorithm that returns numbers and that you want as a string. The maximum base could be extended by adding other characters to the constant DIGITS. With the original DIGITS this algorithm is compatible for conversion to common base: Binary (2), Hex (16), Octal (8).
 
### More Info
 
Use: BaseConv(base10_original_number, newbase)

to convert in a particular base

Or use: ConvBase10(otherbase_number, oldbase)

Juste paste the code anywhere.

BaseConv returns a string and ConvBase10 return a long integer

If you add char to DIGITS (to extend the base maximum) be aware to avoid adding the same char twice or else you will get an error.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vincent Bouret](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vincent-bouret.md)
**Level**          |Beginner
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vincent-bouret-base-conversion-pro__1-11051/archive/master.zip)





### Source Code

```
'Use: BaseConv(base10_original_number, newbase)
'to convert in a particular base
'Or use: ConvBase10(otherbase_number, oldbase)
Const ZERO = "0"
Const DIGITS = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz/.="
Function ConvBase(n1 As Double, base As Long) As String
Dim e2(30) As Double
Dim n As Integer
Dim i As Long
Dim max As Long
If base > Len(DIGITS) Then ConvBase = "NULL": Exit Function
If n1 = 1 Then ConvBase = "1": Exit Function
n = 0
Do While (base ^ n) <= n1
 n = n + 1
Loop
n = n - 1
i = 0
Do While n > -1
 e2(i) = n1 \ (base ^ n)
 n1 = n1 Mod (base ^ n)
 n = n - 1
 i = i + 1
Loop
max = i - 1
For i = 0 To max
 If e2(i) = 0 Then
  ConvBase = ConvBase & ZERO
 Else
  ConvBase = ConvBase & Mid(DIGITS, e2(i), 1)
 End If
Next i
End Function
Function ConvBase10(num As String, base As Long) As Double
Dim i As Long
Dim n As Long
n = Len(num)
For i = 1 To n
 If Mid(num, i, 1) <> ZERO Then
 ConvBase10 = ConvBase10 + (InStr(1, DIGITS, Mid(num, i, 1)) * (base ^ (n - i)))
 End If
Next i
End Function
```

