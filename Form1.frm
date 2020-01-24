VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000015&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   4905
   ClientTop       =   4440
   ClientWidth     =   10005
   FillStyle       =   2  'Horizontal Line
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   6855
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu br 
      Caption         =   "barnameh"
      Begin VB.Menu z 
         Caption         =   "zraeb moadelh"
      End
      Begin VB.Menu zm 
         Caption         =   "zarb motavali"
      End
      Begin VB.Menu n 
         Caption         =   "fact n"
      End
      Begin VB.Menu m 
         Caption         =   "majmoy magsom"
      End
      Begin VB.Menu s 
         Caption         =   "sry fact"
      End
      Begin VB.Menu fe 
         Caption         =   "feobonachi"
      End
      Begin VB.Menu m2 
         Caption         =   "mabna2"
      End
      Begin VB.Menu m20 
         Caption         =   "meangen20"
      End
      Begin VB.Menu j 
         Caption         =   "jadvl zarb"
      End
      Begin VB.Menu mk 
         Caption         =   "meangen karbar"
      End
      Begin VB.Menu ma 
         Caption         =   "majmoy argam"
      End
      Begin VB.Menu v 
         Caption         =   "adad v amalgar"
      End
      Begin VB.Menu bk 
         Caption         =   "bozorg v koch"
      End
      Begin VB.Menu jm 
         Caption         =   "jamy motvali"
      End
      Begin VB.Menu tm 
         Caption         =   "tafreg motvali"
      End
      Begin VB.Menu a3 
         Caption         =   "ary 3,3"
      End
      Begin VB.Menu a2 
         Caption         =   "ary2"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function tafreg(ByRef a As Double, b As Double)
   Dim n As Double
   n = b
   i = 0
   t = 0
   Text1 = " "
   If a = Fix(a) Then
         w = Abs(Fix(a))
         q = Abs(Fix(b))
         While w > q
               w = w - q
               i = i + 1
         Wend
          t = i + (w / q)
          If (a < 0 And b < 0) Or (a > 0 And b > 0) Then Text1 = t Else Text1 = -t
    Else
          MsgBox "no decimal"
    End If
End Function
Private Sub Command1_Click()
Dim a, b, c, x1, x2 As Double
Dim f As Integer
a = InputBox("")
b = InputBox("")
c = InputBox("")
f = (b ^ 2) - (4 * a * c)
If f >= 0 Then
   x1 = (-b + (f ^ (1 / 2))) / (2 * a)
   x2 = (-b - (f ^ (1 / 2))) / (2 * a)
   If x1 = x2 Then
      Print x1, "reshih mozaaf"
   End If
   Print x1, x2
End If

If f < 0 Then Print "no reshih hagegi"
End Sub

Private Sub Command10_Click()
Dim a1, a3, a4, a5 As Double
Dim a2 As Integer
a2 = InputBox("Enter")
If a2 = Abs(a2) Then
   a1 = 1
  
  For a3 = 1 To a2
     a4 = a1 + a5

     Print a4

     a1 = a5

     a5 = a4

  Next
Else
 MsgBox "jhkjh"
  Exit Sub
End If


End Sub

Private Sub Command11_Click()
x = InputBox("")
Dim c(1 To 1000) As Integer

f = 1
While x \ 2
     c(f) = x Mod 2
     x = x \ 2
     f = f + 1
Wend
c(f) = x
Print "(";
For u = f To 1 Step -1
     Print c(u);
Next
Print ")-2"
     
     


End Sub

Private Sub Command12_Click()
Dim i As Integer
Dim x, c As Double
c = 0
For i = 1 To 20
    x = InputBox("")
    c = c + x
Next
Print c / 20
End Sub

Private Sub Command13_Click()
For b = 1 To 10
    For i = 1 To 10
        Print i * b;
    Next
    Print
Next
End Sub

Private Sub Command14_Click()
Dim i, j As Integer
Dim x, c As Double
c = 0
j = InputBox("")
If j = Abs(Fix(j)) Then
  For i = 1 To j
      x = InputBox("")
      c = c + x
  Next
  Print c / j
Else
  Exit Sub
End If
End Sub

Private Sub Command16_Click()
a = InputBox("")
For i = 1 To a
    h = h + i
Next
Print h
End Sub

Private Sub a2_Click()
  Dim a(1 To 10) As Double
  Text1 = " "
  For i = 1 To 10
     a(i) = InputBox("")
  Next
  n = a(1)
  m = a(1)
  For i = 1 To 10
       If n > a(i) Then n = a(i)
       If m < a(i) Then m = a(i)
       b = b + a(i)
  Next
  Text1 = b / 10
  Text1 = Text1 & m
  Text1 = Text & n
End Sub

Private Sub a3_Click()
  Dim a(2, 2) As Double
  Dim b(2, 2) As Double
  Dim i As Integer, m As Integer, t As Integer, k As Integer
  Dim j As Double, n As Double, v As Double, q As Double
  n = 1
  q = 1
  Text1 = " "
  For i = 0 To 2
      For m = 0 To 2
            a(i, m) = InputBox("a" & m)
            j = j + a(i, m)
            n = n * a(i, m)
      Next
  Next
  For t = 0 To 2
      For k = 0 To 2
          b(t, k) = InputBox("b" & k)
          v = v + b(t, k)
          q = q * b(t, k)
      Next
  Next
  Text1 = j & "+" & v & "= " & j + v
  Text1 = Text1 & vbNewLine & n & "*" & q & "= " & n * q
  Text1 = Text1 & vbNewLine & j & "-" & v & "= " & j - v
  Text1 = Text1 & vbNewLine & v & "-" & j & "= " & v - j
End Sub

Private Sub bk_Click()
  Dim f As Double, o As Double, x As Double, i As Integer
  Text1 = " "
  For i = 1 To 5
     x = InputBox("")
     If i = 1 Then
        o = x
        f = x
     End If
     If x > o Then o = x
     If x < f Then f = x
  Next
  Text1 = o & " boz"
  Text1 = Text1 & vbNewLine & f & " koch"
End Sub

Private Sub Command18_Click()
Dim f As Double, o As Double, x As Double, i As Integer
For i = 1 To 5
   x = InputBox("")
  If i = 1 Then o = x
  If i = 1 Then f = x
  
  If x > o Then o = x
  If x < f Then f = x
  Print o; f
Next
Print o; "boz"
Print f; "koch"
End Sub

Private Sub Command19_Click()
Dim a(1 To 10) As Double
For i = 1 To 10
   a(i) = InputBox("")
Next
n = a(1)
m = a(1)
For i = 1 To 10
    If n > a(i) Then n = a(i)
    If m < a(i) Then m = a(i)
    b = b + a(i)
Next
Print b / 10
Print m
Print n
End Sub

Private Sub Command2_Click()
ReDim a(1 To 18) As Double
Dim b, j As Double
b = 123.4
j = b
Print j
Print j * 2
i = 1
While j \ 10 > 10
     j = b
    a(i) = b Mod 10
    b = b \ 10
    
    i = i + 1
Wend
For p = 1 To 18
   Print a(p);
Next
Print b
Print i
End Sub

Private Sub Command20_Click()
Dim a(2, 2) As Double
Dim b(2, 2) As Double
Dim i As Integer, m As Integer, t As Integer, k As Integer
Dim j As Double, n As Double, v As Double, q As Double
n = 1
q = 1
For i = 0 To 2
    For m = 0 To 2
          a(i, m) = InputBox("a" & m)
          j = j + a(i, m)
          n = n * a(i, m)
    Next
Next
For t = 0 To 2
    For k = 0 To 2
        b(t, k) = InputBox("b" & k)
        v = v + b(t, k)
        q = q * b(t, k)
    Next
Next
Print j; "+"; v; "=", j + v
Print n; "*"; q; "=", n * q
Print j; "-"; v; "=", j - v
Print v; "-"; j; "=", v - j
End Sub

Private Sub Command3_Click()
    Dim a, b, j, n, u, t, z, p As Double
    Dim i As Integer
    a = InputBox("pa")
    b = InputBox("tav")
    j = a
   Select Case b + 3 = Fix(b) + 3
     Case True
       Print a; "^"; b; "= ";
       n = Abs(b)
       For i = 1 To n - 1
           j = j * a
           Print a; "*";
       Next
       If b < 0 Then
            u = 1 / j
            Print a; "="; u
       Else
           Print a; "="; j
      End If
  Case False
      p = Len(Str(b)) - 2
      z = "1"
      For t = 1 To p - 1
          z = z + "0"
     Next
    e = CStr(z) * b
    Print a; "^"; b; "="
    Print a ^ (e / CStr(z))
 End Select
End Sub

Private Sub Command4_Click()
Print Abs(Fix(-2.4))
x = InputBox("tavan")
f = InputBox("payeh")
p = Len(Str(x)) - 2
b = "1"
For t = 1 To p - 1
     b = b + "0"
Next
e = CStr(b) * x
Print f ^ (e / CStr(b))
End Sub

Function zarb(ByRef a As Double, b As Double)
  Dim c As Double
  Dim w As Integer
  Text1 = " "
  If a = Fix(a) Then
     w = Abs(Fix(a))
     q = Abs(Fix(b))
     Text1 = Text1 & w
     For i = 1 To w
         c = c + q
     Next
    If (a < 0 And b < 0) Or (a > 0 And b > 0) Then Text1 = c Else Text1 = -c
  Else
     MsgBox "no decimal"
  End If
End Function

Private Sub Command5_Click()
Dim a As Double
Dim b As Double
a = InputBox("")
b = InputBox("")
Call zarb(a, b)





End Sub

Private Sub Command6_Click()
Dim a, b, s As Double
s = 1
a = 0
b = InputBox("")
While b > a
    s = s * b
    b = b - 1
Wend
Print s
End Sub

Private Sub Command7_Click()


Dim i, f As Double

f = 1

For i = 1 To InputBox("Enter a number to reach its single factorial:")

f = f * i

Next

MsgBox f


End Sub

Private Sub Command8_Click()
Dim b, s As Double
b = InputBox("")
For i = 1 To b
    If b Mod i = 0 Then
          Print i;
          s = s + i
    End If
Next
Print "="; s
    
    
End Sub

Private Sub Command9_Click()
Dim p, s, j, i As Double
s = 1
f = InputBox("")
For i = 1 To f
    s = 1
    For j = 1 To i
        s = s * j
    Next
     p = p + (1 / s)
     Print s;
     
Next
Print "="; p + 1

End Sub


Public Function d()
  Dim a As Integer
  Dim b As Integer
  Dim c As Integer
  a = 7
  b = 12
  c = a * b
  ha = c
  
End Function

Private Sub fe_Click()
   Dim a1, a3, a4, a5 As Double
   Dim a2 As Integer
   a2 = InputBox("Enter")
   Text1 = " "
   If a2 + 2 = Abs(Fix(a2)) + 2 Then
      a1 = 1
      For a3 = 1 To a2
        a4 = a1 + a5
        Text1 = Text1 & vbNewLine & a4 & ","
        a1 = a5
        a5 = a4
     Next
  Else
     MsgBox "no manfi and decimal"
     Exit Sub
  End If
End Sub

Private Sub j_Click()
   Text1 = " "
   For b = 1 To 10
      For i = 1 To 10
          Text1 = Text1 & i * b & "   "
          
      Next
      Text1 = Text1 & vbNewLine
   Next
End Sub

Private Sub jm_Click()
  Dim a As Double
  Dim b As Double
  Dim w As Integer
  a = InputBox("1")
  b = InputBox("2")
  Call zarb(a, b)
End Sub

Private Sub m_Click()
    Dim b As Double, s As Double
    s = 0
    b = InputBox("")
    If b = Abs(Fix(b)) Then
       Text1 = " "
       For i = 1 To b
           If b Mod i = 0 Then
               Text1 = Text1 & i & ","
                s = s + i
           End If
       Next
       Text1 = Text1 & "=" & s
   Else
       MsgBox "no decimal and manfi"
   End If
End Sub

Private Sub m2_Click()
Dim x As Integer
   x = InputBox("")
   Text1 = " "
   If x = Abs(Fix(x)) Then
      Dim c(1 To 1000) As Integer
      f = 1
      While x \ 2
         c(f) = x Mod 2
         x = x \ 2
         f = f + 1
     Wend
     c(f) = x
     Text1 = "("
     For u = f To 1 Step -1
            Text1 = Text1 & c(u)
     Next
     Text1 = Text1 & ")_2"
  Else
      MsgBox "no manfi and decimal"
  End If
End Sub

Private Sub m20_Click()
  Dim i As Integer
  Dim x As Double, c As Double
  c = 0
  Text1 = " "
  For i = 1 To 20
    x = InputBox("")
    c = c + x
  Next
  Text1 = Text1 & c / 20
End Sub

Private Sub ma_Click()
   Dim a As Double, h As Double
   a = InputBox("")
   h = 0
   Text1 = " "
  If a = Fix(a) Then
     w = Abs(a)
     For i = 1 To w
         h = h + i
     Next
     If a < 0 Then Text1 = Text & -h Else Text1 = Text & h
    
  Else
     MsgBox "no decimal"
  End If
End Sub

Private Sub mk_Click()
  Dim i As Integer
   Dim j As Double
  Dim x As Double, c As Double
  c = 0
  j = InputBox("")
  Text1 = " "
  If j + 2 = Abs(Fix(j)) + 2 Then
    For i = 1 To j
        x = InputBox("")
        c = c + x
    Next
    Text1 = c / j
  Else
   MsgBox "no decimal and manfi"
  End If
End Sub

Private Sub n_Click()
    Dim a As Double, b As Double, s As Double
    s = 1
    a = 0
    b = InputBox("")
    Text1 = " "
    If b + 3 = Fix(b) + 3 Then
        u = Abs(b)
       While u > a
            s = s * u
            u = u - 1
       Wend
       If b < 0 Then Text1 = b & "!" & "=" & -s Else Text1 = b & "!" & "=" & s
   Else
       MsgBox "no decimal"
   End If
End Sub

Private Sub s_Click()
  Dim p As Double, s As Double, j As Double, i As Double
  s = 1
  f = InputBox("")
  Text1 = " "
  If f + 3 = Fix(f) + 3 Then
     y = Abs(f)
     For i = 1 To y
          s = 1
          For j = 1 To i
          s = s * j
     Next
     p = p + (1 / s)
     If f < 0 Then Text1 = Text1 & -p & " + " Else Text1 = Text1 & p & " + "
     Next
    If f < 0 Then Text1 = Text1 & "= " & -p + 1 Else Text1 = Text1 & "= " & p + 1
  Else
     MsgBox "no decimal"
  End If
End Sub

Private Sub tm_Click()
  Dim a As Double, b As Double
  a = InputBox("1")
  b = InputBox("2")
  Call tafreg(a, b)
End Sub

Private Sub v_Click()
   a = InputBox("1")
   b = InputBox("2")
   c = InputBox("amalgar")
   Text1 = " "
  Select Case c
       Case "*"
         Text1 = b * a
       Case "="
         Print a = b
       Case "/"
           Text1 = Text1 & a / b
           Text1 = Text & b / a
      Case "\"
          Text1 = txt1 & b \ a
          Text1 = Text1 & a \ b
      Case "+"
          Text1 = a + b
      Case "-"
          Text1 = Text1 & a - b
          Text1 = Text1 & b - a
      Case ">"
          Text1 = Text1 & a > b
          Text1 = Text1 & b > a
      Case "<"
          Text1 = Text1 & a < b
          Text1 = Text1 & b < a
      Case "<>"
           Text1 = a <> b
      Case "^"
           Text1 = Text1 & a ^ b
           Text1 = Text1 & b ^ a
      Case ">="
           Text1 = Text1 & a >= b
           Text1 = Text1 & b >= a
      Case "<="
           Text1 = Text1 & a <= b
           Text1 = Text1 & b <= a
      Case "mod"
           Text1 = Text1 & a Mod b
           Text1 = Text1 & b Mod a
End Select
End Sub

Private Sub z_Click()
     Dim a As Double, b As Double, c As Double, x1 As Double, x2 As Double
     Dim f As Double
     a = InputBox("a")
     b = InputBox("b")
     c = InputBox("c")
     f = (b ^ 2) - (4 * a * c)
     Text1 = " "
     If f >= 0 Then
        x1 = (-b + (f ^ (1 / 2))) / (2 * a)
        x2 = (-b - (f ^ (1 / 2))) / (2 * a)
        If x1 = x2 Then
             Text1 = "resheh Mozaf" & vbNewLine & "x1=  " & x1
        Else
             Text1 = "reshehHaghighi" & vbNewLine & "x1=  " & x1 & vbNewLine & "x2=  " & x2
        End If
    End If
    If f < 0 Then Text1 = "no resheh "
End Sub

Private Sub zm_Click()
    Dim a As Double, b As Double, j As Double, n As Double, u As Double, t As Double, z As Double, p As Double
    Dim i As Integer
    Dim e As Double, lq As Double
    a = InputBox("pa")
    b = InputBox("tav")
    j = 1
    Text1 = ""
   Select Case b + 3 = Fix(b) + 3
     Case True
       Text1 = a & "^" & b & "= "
       n = Abs(b)
       For i = 1 To n
           j = j * a
           If i > 1 Then Text1 = Text1 & a & "*"
           
       Next
       Text1 = Text1 & a
       If b < 0 Then
            u = 1 / j
            Text1 = Text1 & "=  " & u
       Else
           Text1 = Text1 & "=  " & j
      End If
  Case False
      p = Len(Str(b)) - 2
      z = "1"
      For t = 1 To p - 1
          z = z + "0"
     Next
    e = CStr(z) * b
    lq = CStr(z)
    Text1 = a & "^" & b & "="
    Text1 = Text1 & a ^ (e / lq)
 End Select
End Sub
