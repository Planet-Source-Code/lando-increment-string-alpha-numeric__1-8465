<div align="center">

## INCREMENT STRING\(ALPHA/NUMERIC\)


</div>

### Description

Increments text string with alpha and numeric characters.
 
### More Info
 
text string

Add two text boxes and a command button to a form.

Returns incremented text


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LANDO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lando.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lando-increment-string-alpha-numeric__1-8465/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
thetext = Text1.Text
incremnttxt = findchr(thetext)
Text2.Text = incremnttxt
End Sub
Function findchr(ByVal thetext As String)
Dim strlen As Integer
Dim A1() As String
strlen = Len(thetext)   ' number of characters
ReDim A1(strlen)
For L = 1 To UBound(A1)  ' parse individual characters
A1(L) = Mid(thetext, L, 1)
Next L
For nxtchar = 1 To UBound(A1)  ' cyle through characters increment ascii value
valchar = (UBound(A1)) - (nxtchar - 1)
 If Asc(A1(valchar)) >= 65 And Asc(A1(valchar)) <= 90 Or _
 Asc(A1(valchar)) >= 97 And Asc(A1(valchar)) <= 122 Then  ' upper and lower alpha characters
  If Asc(A1(valchar)) = 90 Or Asc(A1(valchar)) = 122 Then
   If Asc(A1(valchar)) = 90 Then
    If valchar = 1 Then ' fisrt char at the end of ascii list
    A1(valchar) = "AA"
    Else
    A1(valchar) = "A"
    End If
   Else
    If valchar = 1 Then ' fisrt char at the end of ascii list
    A1(valchar) = "aa"
    Else
    A1(valchar) = "a"
    End If
   End If
  Else
  A1(valchar) = Chr(Asc(A1(valchar)) + 1) ' increment ascii by one
  GoTo noneedto:
  End If
 ElseIf Asc(A1(valchar)) > 47 And Asc(A1(valchar)) < 58 Then 'numeric values
   If Asc(A1(valchar)) = 57 Then
    If valchar = 1 Then ' fisrt char at the end of ascii list
    A1(valchar) = "10"
    Else
    A1(valchar) = "0"
    End If
   Else
   A1(valchar) = Chr(Asc(A1(valchar)) + 1) ' increment ascii by one
   GoTo noneedto:
   End If
 End If
Next nxtchar
noneedto: 'once a char is increment and is not carried over no need to increment all chars
For mke = LBound(A1) To UBound(A1) ' make text
findchr = Trim$(findchr) & A1(mke)
Next mke
End Function
```

