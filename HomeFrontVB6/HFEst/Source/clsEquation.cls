VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const vbQuote = """"

Private Dirty               As Boolean
Private Parsed              As Boolean
Private Deg                 As Boolean
Private Equ                 As String
Private dAnswer             As Double
Private EquParsed           As Collection
Private EquOrder            As Collection
Private Vars                As Collection

Private Enum Priorities
     PRI_None
     PRI_BOOLEAN
     PRI_OPERATOR
     PRI_ADD
     PRI_MOD
     PRI_MUL
     PRI_NEG
     PRI_EXP
     PRI_VAR
     PRI_PAR
     PRI_LEVEL = PRI_PAR
End Enum


Private Const EQ_NONE       As Integer = 0
Private Const EQ_STRING     As Integer = 1
Private Const EQ_NUMBER     As Integer = 2

Private Const ER_NONE       As Integer = 0
Private Const ER_VAR        As Integer = 1

Private Const EQ_PAREN      As Integer = 1100
Private Const EQ_FUNCTION   As Integer = 1101
Private Const EQ_VARIABLE   As Integer = 1102
Private Const EQ_INVALID    As Integer = 1103
Private Const EQ_ARGS       As Integer = 1104
Private Const EQ_NAME       As Integer = 1105


Private Const PI            As Double = 3.14159265358979
Private Const DEG_TO_RAD    As Double = PI / 180    ''0.01745329251995
Private Const RAD_TO_DEG    As Double = 180 / PI    ''57.2957795131

Public ParentForm As Object

Private Sub Class_Initialize()
    
    Set EquParsed = New Collection
    Set EquOrder = New Collection
    Set Vars = New Collection
    
    Dirty = False
    Parsed = True
    Degrees = False
    
End Sub

Private Sub Class_Terminate()

    Set EquParsed = Nothing
    Set EquOrder = Nothing
    Set Vars = Nothing

End Sub


Public Property Let Degrees(b As Boolean)
    If b <> Deg Then
        Deg = b
        Dirty = True
    End If
End Property

Public Property Get Degrees() As Boolean
    Degrees = Deg
End Property

Public Property Let Equation(E As String)
    Parsed = False
    Dirty = True
'    Equ = Replace(Replace(E, vbLf, " "), vbCr, " ")
' added +0 to end of equation to resolve a parsing error. Easiest solution :)
     Equ = Replace(Replace("(" & E & ")+0", vbLf, " "), vbCr, " ")
    
End Property

Public Property Get Equation() As String
    Equation = Equ
End Property

Public Property Get Var(Name As String) As Double
    On Local Error GoTo GetError
    Var = CDbl(Vars(Name))
    Exit Property
GetError:
    Var = 0#
End Property

Public Property Let Var(Name As String, Num As Double)
    On Local Error Resume Next
    Dirty = True
    Vars.Remove Trim$(Name)
    Vars.Add Num, Trim$(Name)
End Property

Public Sub VarClear()
    Set Vars = New Collection
    Dirty = True
End Sub

Public Sub VarRemove(Name As String)
    On Local Error Resume Next
    Vars.Remove Name
    Dirty = True
End Sub

Public Function Solution() As Double
    If Dirty Then Call Solve
    Solution = dAnswer
End Function
Public Sub Solve()
Dim z As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim i1                  As Long
    Dim i2                  As Long
    Dim i3                  As Long
    Dim i4                  As Long
    Dim i5                  As Long
    Dim i6                  As Long
    
    Dim m                   As Long
    Dim n                   As Long
    
    Dim P1                  As Double
    Dim P2                  As Double
    Dim P3                  As Double
    Dim P4                  As Double
    Dim p5                  As Double
    Dim p6                  As Double
    
    Dim s1                  As String
    Dim s2                  As String
    Dim s3                  As String
    Dim s4                  As String
    Dim s5                  As String
    Dim s6                  As String
    
    Dim iParams             As Long
    Dim v1                  As Variant
    Dim v2                  As Variant
        
    Dim v                   As Variant
    Dim eSpace              As Integer
    Dim Temp()              As Variant
    Dim j2                  As Long
    
    On Local Error GoTo SolveError
    
    If Not Parsed Then Parse

    ReDim Temp(0 To EquParsed.Count)
    For i = 1 To EquParsed.Count
        Temp(i) = EquParsed(i)
    Next
    eSpace = ER_NONE
    For i = 1 To EquOrder.Count
    
        m = EquOrder(i)
        v = Temp(m)
        i1 = GetLeft(m, Temp)
        i2 = GetRight(m, Temp)
        
        
        If i1 > 0 And i1 <= UBound(Temp) Then v1 = Temp(i1)
        If i2 > 0 And i2 <= UBound(Temp) Then v2 = Temp(i2)
        
        If IsNumeric(v1) And IsNumeric(v2) Then
            v1 = CDbl(v1)
            v2 = CDbl(v2)
        End If
            
        Select Case v
            Case "~":    Temp(m) = "" & -CDbl(Temp(i2))
                         Temp(i2) = Null
            Case "=":    Temp(i1) = "" & IIf(v1 = v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "<":    Temp(i1) = "" & IIf(v1 < v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case ">":    Temp(i1) = "" & IIf(v1 > v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(2): Temp(i1) = "" & IIf(v1 <= v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(3): Temp(i1) = "" & IIf(v1 >= v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(4): Temp(i1) = "" & IIf(v1 <> v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case Chr(5): Temp(i1) = "" & IIf(v1 And v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case Chr(6): Temp(i1) = "" & IIf(v1 Or v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                          
                         
            Case "+"
                         If IsNumeric(Temp(i1)) Then
                             Temp(i1) = "" & CDbl(Temp(i1)) + CDbl(Temp(i2))
                         Else
                             Temp(i1) = Quote(UnQuote(Temp(i1)) + UnQuote(Temp(i2)))
                         End If
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case "*":    Temp(i1) = "" & CDbl(Temp(i1)) * CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "/":    Temp(i1) = "" & CDbl(Temp(i1)) / CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "\":    Temp(i1) = "" & CDbl(Temp(i1)) \ CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "-":    Temp(i1) = "" & CDbl(Temp(i1)) - CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "^":    Temp(i1) = "" & CDbl(Temp(i1)) ^ CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "%":    Temp(i1) = "" & CDbl(Temp(i1)) Mod CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
                
            Case "("
                i = i + 1
                n = EquOrder(i)
                i2 = GetRight(m, Temp)
                If i2 >= n Then
                    Temp(m) = 0#
                    Temp(n) = Null
                Else
                    Temp(m) = Temp(i2)
                    Temp(i2) = Null
                    Temp(n) = Null
                End If
            Case Else
                If Right$(Temp(m), 1) = "(" Then
                    i1 = 0
                    i2 = 0
                    i3 = 0
                    i4 = 0
                    i5 = 0
                    i6 = 0
                    s1 = ""
                    s2 = ""
                    s3 = ""
                    s4 = ""
                    s5 = ""
                    s6 = ""
                    P1 = 0
                    P2 = 0
                    P3 = 0
                    P4 = 0
                    p5 = 0
                    p6 = 0
                    
                    
                    i = i + 1
                    n = EquOrder(i)
                    
                    i1 = GetRight(m, Temp)
                    i2 = GetRight(i1, Temp)
                    i3 = GetRight(i2, Temp)
                    i4 = GetRight(i3, Temp)
                    i5 = GetRight(i4, Temp)
                    'i6 = GetRight(i5, Temp)
                    i6 = GetLeft(n, Temp)
                    
                    iParams = 6
                    If Temp(i6) = ")" Then iParams = 5
                    If Temp(i5) = ")" Then iParams = 4
                    If Temp(i4) = ")" Then iParams = 3
                    If Temp(i3) = ")" Then iParams = 2
                    If Temp(i2) = ")" Then iParams = 1
                    If Temp(i1) = ")" Then iParams = 0

                    If i1 >= n Then i1 = 0
                    P1 = Val("" & Temp(i1))
                    s1 = "" & Temp(i1)
                       
                    If i2 <= i1 Then i2 = 0
                    P2 = Val("" & Temp(i2))
                    s2 = "" & Temp(i2)
                                        
                    If i3 > n Then i3 = 0
                    P3 = Val("" & Temp(i3))
                    s3 = "" & Temp(i3)
                                        
                    If i4 > n Then i4 = 0
                    P4 = Val("" & Temp(i4))
                    s4 = "" & Temp(i4)
                                        
                    If i5 > n Then i5 = 0
                    p5 = Val("" & Temp(i5))
                    s5 = "" & Temp(i5)
                                        
                    If i6 > n Then i6 = 0
                    p6 = Val("" & Temp(i6))
                    s6 = "" & Temp(i6)
                                        
                    s1 = UnQuote(s1)
                    s2 = UnQuote(s2)
                    s3 = UnQuote(s3)
                    s4 = UnQuote(s4)
                    s5 = UnQuote(s5)
                    s6 = UnQuote(s6)
                                        
                    Temp(i1) = Null
                    Temp(i2) = Null
                    Temp(i3) = Null
                    Temp(i4) = Null
                    Temp(i5) = Null
                    Temp(i6) = Null
                    
                    Temp(m) = Null
                    Temp(n) = Null
                    Select Case v
                        Case "abs(":       Temp(m) = Abs(P1)
                        Case "arccos(":    Temp(m) = (Atn(-P1 / Sqr(-P1 * P1 + 1)) + 2 * Atn(1)) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccosec(":  Temp(m) = (Atn(P1 / Sqr(P1 * P1 - 1)) + (Sgn(P1) - 1) * (2 * Atn(1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccosech(": Temp(m) = Log((Sgn(P1) * Sqr(P1 * P1 + 1) + 1) / P1)
                        Case "arccosh(":   Temp(m) = Log(P1 + Sqr(P1 * P1 - 1))
                        Case "arccotan(":  Temp(m) = (Atn(P1 * IIf(Degrees, RAD_TO_DEG, 1)) + 2 * Atn(1)) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccotanh(": Temp(m) = Log((P1 + 1) / (P1 - 1)) / 2
                        Case "arcsec(":    Temp(m) = (Atn(P1 / Sqr(P1 * P1 - 1)) + (Sgn(P1) - 1) * (2 * Atn(1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arcsech(":   Temp(m) = Log((Sqr(-P1 * P1 + 1) + 1) / P1)
                        Case "arcsin(":    Temp(m) = (Atn(P1 / Sqr(-P1 * P1 + 1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arcsinh(":   Temp(m) = Log(P1 + Sqr(P1 * P1 + 1))
                        Case "arctan(":    Temp(m) = Atn(P1) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arctanh(":   Temp(m) = Log((1 + P1) / (1 - P1)) / 2
                        Case "atn(":       Temp(m) = Atn(P1) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "avg(":       Temp(m) = Avg(P1, P2, P3)
                        Case "ceiling(":   Temp(m) = Ceiling(P1)
                        Case "cos(":       Temp(m) = Cos(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "cosec(":     Temp(m) = (1 / Sin(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "cosech(":    Temp(m) = 2 / (Exp(P1) - Exp(-P1))
                        Case "cosh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / (Exp(P1) + Exp(-P1))
                        Case "cot(":       Temp(m) = (1 / Tan(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "cotanh(":    Temp(m) = (Exp(P1) + Exp(-P1)) / (Exp(P1) - Exp(-P1))
                        Case "degrees(":   Temp(m) = P1 * RAD_TO_DEG
                        Case "exp(":       Temp(m) = Exp(P1)
                        Case "fix(":       Temp(m) = Fix(P1)
                        Case "floor(":     Temp(m) = Floor(P1)
                        Case "int(":       Temp(m) = Int(P1)
                        Case "ln(":        Temp(m) = Log(P1)
                        Case "log(":       Temp(m) = Log(P1)
                        Case "log10(":     Temp(m) = Log(P1) / Log(10)
                        Case "log2(":      Temp(m) = Log(P1) / Log(2)
                        Case "logn(":      Temp(m) = Log(P1) / Log(P2)
                        Case "max(":       Temp(m) = Max(P1, P2, P3)
                        Case "min(":       Temp(m) = Min(P1, P2, P3)
                        Case "mod(":       Temp(m) = P1 Mod P2
                        Case "pi(":        Temp(m) = PI
                        Case "power(":     Temp(m) = P1 ^ P2
                        Case "radians(":   Temp(m) = P1 * DEG_TO_RAD
                        Case "rnd(":       Temp(m) = Rnd
                        Case "rand(":      Temp(m) = Int(Rnd * P1)
                        Case "random(":    Temp(m) = (Rnd * (P2 - P1)) + P1
                        Case "round(":     Temp(m) = Round(P1, P2)
                        Case "roundto(":   Temp(m) = RoundTo(P1, P2, P3)
                        Case "sec(":       Temp(m) = (1 / Cos(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "sech(":      Temp(m) = 2 / (Exp(P1) + Exp(-P1))
                        Case "sgn(":       Temp(m) = Sgn(P1)
                        Case "sin(":       Temp(m) = Sin(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "sinh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / 2
                        Case "sqrt(":      Temp(m) = Sqr(P1)
                        Case "square(":    Temp(m) = P1 * P1
                        Case "sum(":       Temp(m) = sum(P1, P2, P3)
                        Case "tan(":       Temp(m) = Tan(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "tanh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / (Exp(P1) + Exp(-P1))
                        Case "trunc(":     Temp(m) = trunc(P1, P2)
                        
                        'local sql function...
                        Case "assemblyorderqty(":     Call CheckArgs(v, iParams, 6):  Temp(m) = GetAssemblyQty("OrderQty", s1, s2, s3, s4, s5, s6)
                        Case "assemblytakeoffqty(":   Call CheckArgs(v, iParams, 6):  Temp(m) = GetAssemblyQty("TakeoffQty", s1, s2, s3, s4, s5, s6)
                        
                        'if you add any new lookup functions you must implement them in FTakeoff and FFormulaTest
                        Case "totalorderqty(":   Call CheckArgs(v, iParams, 4):  Temp(m) = ParentForm.GetItemTotal("OrderQty", s1, s2, s3, s4)
                        Case "totaltakeoffqty(": Call CheckArgs(v, iParams, 4):  Temp(m) = ParentForm.GetItemTotal("TakeoffQty", s1, s2, s3, s4)
                        Case "totalcost(":       Call CheckArgs(v, iParams, 4):  Temp(m) = ParentForm.GetItemTotal("Cost", s1, s2, s3, s4)
                        
                        Case "if("
                            If IsNumeric(s2) Or IsNumeric(s3) Then
                                Temp(m) = IIf(P1, P2, P3)
                            Else
                                Temp(m) = Quote(IIf(P1, s2, s3))
                            End If
                        
                        Case "char(":      Temp(m) = Quote(Chr(P1))
                        Case "charindex(": Temp(m) = InStr(1, s1, s2)
                        Case "instr(":     Temp(m) = InStr(1, s1, s2)
                        Case "left(":      Temp(m) = Quote(Left(s1, P2))
                        Case "len(":       Temp(m) = Len(s1)
                        Case "lower(":     Temp(m) = Quote(LCase(s1))
                        Case "ltrim(":     Temp(m) = Quote(LTrim(s1))
                        Case "replace(":   Temp(m) = Quote(Replace(s1, s2, s3))
                        Case "replicate(": Temp(m) = Quote(String(P2, s1))
                        Case "right(":     Temp(m) = Quote(Right(s1, P2))
                        Case "rtrim(":     Temp(m) = Quote(RTrim(s1))
                        Case "space(":     Temp(m) = Quote(String(P1, " "))
                        Case "str(":       Temp(m) = Quote("" & P1)
                        Case "substring(": Temp(m) = Quote(Mid(s1, P2, P3))
                        Case "upper(":     Temp(m) = Quote(UCase(s1))
                        
                        Case Else
                            Err.Raise EQ_FUNCTION, "clsEquation", "Undefined Function: " & v & ")"
                            Exit Sub
                    End Select
                Else
                    Select Case v
                        Case "pi":       Temp(m) = PI
                        Case "e":        Temp(m) = 2.718281828
                        Case "rnd":      Temp(m) = Rnd
                        Case Else
                            If Left(v, 1) = """" Then
                                Temp(m) = v
                            Else
                                eSpace = ER_VAR
                                Temp(m) = CDbl(Vars(Temp(m)))
                                eSpace = ER_NONE
                            End If
                    End Select
                End If
        End Select
    Next i
   
    dAnswer = CDbl(Temp(GetRight(0, Temp)))
    Dirty = False
    Exit Sub

SolveError:
    Select Case Err
        Case 6, 11, EQ_PAREN To EQ_NAME
            Err.Raise Err, "clsEquation", Err.Description
        Case 5:
            Select Case eSpace
                Case ER_VAR
                    Err.Raise EQ_VARIABLE, "clsEquation", "Undefined Function: " & v
                Case Else
                    Err.Raise Err, "clsEquation", Err.Description
            End Select
        Case Else
            Err.Raise EQ_INVALID, "clsEquation", "Invalid Equation"
    End Select
End Sub


Private Sub CheckArgs(FunctionName As Variant, Given As Long, Expected As Long)
    If Given <> Expected Then
        Err.Raise EQ_ARGS, "clsEquation", Replace(FunctionName, "(", "") & " expects " & Expected & " argument(s)"
    End If
End Sub


Private Function GetAssemblyQty(QtyType As String, community As String, Assembly As String, Model As String, OptionID As String, Phase As String, Item As String) As Double
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim qty As Double
    
    If Item = "" Or Phase = "" Or Assembly = "" Then
        Call Err.Raise(2, "GetAssemblyQty", "Invalid parameters." & vbCrLf & "Expected: community,assembly,model,optionid,phase,item" & vbCrLf & vbCrLf & "Parameters assembly, phase, item may not be blank.")
    End If
    
    qty = 0
    
    s = ""
    s = s & "SELECT dbo.Purch_GetAssemblyQty(" & DbQuote(Num, HFApp.DivisionID) & ","
    s = s & DbQuote(Str, QtyType) & ","
    s = s & DbQuote(Str, community) & ","
    s = s & DbQuote(Str, Assembly) & ","
    s = s & DbQuote(Str, Model) & ","
    s = s & DbQuote(Str, OptionID) & ","
    s = s & DbQuote(Str, Phase) & ","
    s = s & DbQuote(Str, Item) & ")"
    
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then qty = Val("" & rs(0))
                    
    GetAssemblyQty = qty
    
Exit Function
eh: MsgBox Err.Description, vbExclamation, "Unable to process function"
End Function

Public Sub SolveOLD()
Dim z As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim i1                  As Long
    Dim i2                  As Long
    Dim i3                  As Long
    Dim i4                  As Long
    Dim m                   As Long
    Dim n                   As Long
    
    Dim P1                  As Double
    Dim P2                  As Double
    Dim P3                  As Double
    Dim P4                  As Double
    Dim s1                  As String
    Dim s2                  As String
    Dim s3                  As String
    Dim s4                  As String
    Dim v1                  As Variant
    Dim v2                  As Variant
        
    Dim v                   As Variant
    Dim eSpace              As Integer
    Dim Temp()              As Variant
    Dim j2                  As Long
    
    On Local Error GoTo SolveError
    
    If Not Parsed Then Parse

    ReDim Temp(0 To EquParsed.Count)
    For i = 1 To EquParsed.Count
        Temp(i) = EquParsed(i)
    Next
    eSpace = ER_NONE
    For i = 1 To EquOrder.Count
    
        m = EquOrder(i)
        v = Temp(m)
        i1 = GetLeft(m, Temp)
        i2 = GetRight(m, Temp)
        
        
        If i1 > 0 And i1 <= UBound(Temp) Then v1 = Temp(i1)
        If i2 > 0 And i2 <= UBound(Temp) Then v2 = Temp(i2)
        
        If IsNumeric(v1) And IsNumeric(v2) Then
            v1 = CDbl(v1)
            v2 = CDbl(v2)
        End If
            
        Select Case v
            Case "~":    Temp(m) = "" & -CDbl(Temp(i2))
                         Temp(i2) = Null
            Case "=":    Temp(i1) = "" & IIf(v1 = v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "<":    Temp(i1) = "" & IIf(v1 < v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case ">":    Temp(i1) = "" & IIf(v1 > v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(2): Temp(i1) = "" & IIf(v1 <= v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(3): Temp(i1) = "" & IIf(v1 >= v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
            Case Chr(4): Temp(i1) = "" & IIf(v1 <> v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case Chr(5): Temp(i1) = "" & IIf(v1 And v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case Chr(6): Temp(i1) = "" & IIf(v1 Or v2, 1, 0)
                         Temp(i2) = Null
                         Temp(m) = Null
                          
                         
            Case "+"
                         If IsNumeric(Temp(i1)) Then
                             Temp(i1) = "" & CDbl(Temp(i1)) + CDbl(Temp(i2))
                         Else
                             Temp(i1) = Quote(UnQuote(Temp(i1)) + UnQuote(Temp(i2)))
                         End If
                         Temp(i2) = Null
                         Temp(m) = Null
                         
            Case "*":    Temp(i1) = "" & CDbl(Temp(i1)) * CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "/":    Temp(i1) = "" & CDbl(Temp(i1)) / CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "\":    Temp(i1) = "" & CDbl(Temp(i1)) \ CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "-":    Temp(i1) = "" & CDbl(Temp(i1)) - CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "^":    Temp(i1) = "" & CDbl(Temp(i1)) ^ CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
            Case "%":    Temp(i1) = "" & CDbl(Temp(i1)) Mod CDbl(Temp(i2))
                         Temp(i2) = Null
                         Temp(m) = Null
                
            Case "("
                i = i + 1
                n = EquOrder(i)
                i2 = GetRight(m, Temp)
                If i2 >= n Then
                    Temp(m) = 0#
                    Temp(n) = Null
                Else
                    Temp(m) = Temp(i2)
                    Temp(i2) = Null
                    Temp(n) = Null
                End If
            Case Else
                If Right$(Temp(m), 1) = "(" Then
                    i = i + 1
                    n = EquOrder(i)
                    
                    i1 = GetRight(m, Temp)
                    i2 = GetRight(i1, Temp)
                    i3 = GetRight(i2, Temp)
                    i4 = GetLeft(n, Temp)
                    
                    If i1 >= n Then i1 = 0
                    P1 = Val(Temp(i1))
                    s1 = Temp(i1)
                       
                    If i2 <= i1 Then i2 = 0
                    P2 = Val(Temp(i2))
                    s2 = Temp(i2)
                                        
                    If i3 > n Then i3 = 0
                    P3 = Val(Temp(i3))
                    s3 = Temp(i3)
                                        
                    If i4 > n Then i4 = 0
                    P4 = Val(Temp(i4))
                    s4 = Temp(i4)
                                        
                    s1 = UnQuote(s1)
                    s2 = UnQuote(s2)
                    s3 = UnQuote(s3)
                    s4 = UnQuote(s4)
                                        
                    Temp(i1) = Null
                    Temp(i2) = Null
                    Temp(i3) = Null
                    Temp(i4) = Null
                    Temp(m) = Null
                    Temp(n) = Null
                    Select Case v
                        Case "abs(":       Temp(m) = Abs(P1)
                        Case "arccos(":    Temp(m) = (Atn(-P1 / Sqr(-P1 * P1 + 1)) + 2 * Atn(1)) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccosec(":  Temp(m) = (Atn(P1 / Sqr(P1 * P1 - 1)) + (Sgn(P1) - 1) * (2 * Atn(1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccosech(": Temp(m) = Log((Sgn(P1) * Sqr(P1 * P1 + 1) + 1) / P1)
                        Case "arccosh(":   Temp(m) = Log(P1 + Sqr(P1 * P1 - 1))
                        Case "arccotan(":  Temp(m) = (Atn(P1 * IIf(Degrees, RAD_TO_DEG, 1)) + 2 * Atn(1)) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arccotanh(": Temp(m) = Log((P1 + 1) / (P1 - 1)) / 2
                        Case "arcsec(":    Temp(m) = (Atn(P1 / Sqr(P1 * P1 - 1)) + (Sgn(P1) - 1) * (2 * Atn(1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arcsech(":   Temp(m) = Log((Sqr(-P1 * P1 + 1) + 1) / P1)
                        Case "arcsin(":    Temp(m) = (Atn(P1 / Sqr(-P1 * P1 + 1))) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arcsinh(":   Temp(m) = Log(P1 + Sqr(P1 * P1 + 1))
                        Case "arctan(":    Temp(m) = Atn(P1) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "arctanh(":   Temp(m) = Log((1 + P1) / (1 - P1)) / 2
                        Case "atn(":       Temp(m) = Atn(P1) * IIf(Degrees, RAD_TO_DEG, 1)
                        Case "avg(":       Temp(m) = Avg(P1, P2, P3)
                        Case "ceiling(":   Temp(m) = Ceiling(P1)
                        Case "cos(":       Temp(m) = Cos(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "cosec(":     Temp(m) = (1 / Sin(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "cosech(":    Temp(m) = 2 / (Exp(P1) - Exp(-P1))
                        Case "cosh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / (Exp(P1) + Exp(-P1))
                        Case "cot(":       Temp(m) = (1 / Tan(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "cotanh(":    Temp(m) = (Exp(P1) + Exp(-P1)) / (Exp(P1) - Exp(-P1))
                        Case "degrees(":   Temp(m) = P1 * RAD_TO_DEG
                        Case "exp(":       Temp(m) = Exp(P1)
                        Case "fix(":       Temp(m) = Fix(P1)
                        Case "floor(":     Temp(m) = Floor(P1)
                        Case "int(":       Temp(m) = Int(P1)
                        Case "ln(":        Temp(m) = Log(P1)
                        Case "log(":       Temp(m) = Log(P1)
                        Case "log10(":     Temp(m) = Log(P1) / Log(10)
                        Case "log2(":      Temp(m) = Log(P1) / Log(2)
                        Case "logn(":      Temp(m) = Log(P1) / Log(P2)
                        Case "max(":       Temp(m) = Max(P1, P2, P3)
                        Case "min(":       Temp(m) = Min(P1, P2, P3)
                        Case "mod(":       Temp(m) = P1 Mod P2
                        Case "pi(":        Temp(m) = PI
                        Case "power(":     Temp(m) = P1 ^ P2
                        Case "radians(":   Temp(m) = P1 * DEG_TO_RAD
                        Case "rnd(":       Temp(m) = Rnd
                        Case "rand(":      Temp(m) = Int(Rnd * P1)
                        Case "random(":    Temp(m) = (Rnd * (P2 - P1)) + P1
                        Case "round(":     Temp(m) = Round(P1, P2)
                        Case "roundto(":   Temp(m) = RoundTo(P1, P2, P3)
                        Case "sec(":       Temp(m) = (1 / Cos(P1 * IIf(Degrees, RAD_TO_DEG, 1)))
                        Case "sech(":      Temp(m) = 2 / (Exp(P1) + Exp(-P1))
                        Case "sgn(":       Temp(m) = Sgn(P1)
                        Case "sin(":       Temp(m) = Sin(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "sinh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / 2
                        Case "sqrt(":      Temp(m) = Sqr(P1)
                        Case "square(":    Temp(m) = P1 * P1
                        Case "sum(":       Temp(m) = sum(P1, P2, P3)
                        Case "tan(":       Temp(m) = Tan(P1 * IIf(Degrees, RAD_TO_DEG, 1))
                        Case "tanh(":      Temp(m) = (Exp(P1) - Exp(-P1)) / (Exp(P1) + Exp(-P1))
                        Case "trunc(":     Temp(m) = trunc(P1, P2)
                        
                        'if you add any new lookup functions you must implement them in FTakeoff and FFormulaTest
                        Case "totalorderqty(":   Temp(m) = ParentForm.GetItemTotal("OrderQty", s1, s2, s3, s4)
                        Case "totaltakeoffqty(": Temp(m) = ParentForm.GetItemTotal("TakeoffQty", s1, s2, s3, s4)
                        Case "totalcost(":       Temp(m) = ParentForm.GetItemTotal("Cost", s1, s2, s3, s4)
                        
                        Case "if("
                            If IsNumeric(s2) Or IsNumeric(s3) Then
                                Temp(m) = IIf(P1, P2, P3)
                            Else
                                Temp(m) = Quote(IIf(P1, s2, s3))
                            End If
                        
                        Case "char(":      Temp(m) = Quote(Chr(P1))
                        Case "charindex(": Temp(m) = InStr(1, s1, s2)
                        Case "instr(":     Temp(m) = InStr(1, s1, s2)
                        Case "left(":      Temp(m) = Quote(Left(s1, P2))
                        Case "len(":       Temp(m) = Len(s1)
                        Case "lower(":     Temp(m) = Quote(LCase(s1))
                        Case "ltrim(":     Temp(m) = Quote(LTrim(s1))
                        Case "replace(":   Temp(m) = Quote(Replace(s1, s2, s3))
                        Case "replicate(": Temp(m) = Quote(String(P2, s1))
                        Case "right(":     Temp(m) = Quote(Right(s1, P2))
                        Case "rtrim(":     Temp(m) = Quote(RTrim(s1))
                        Case "space(":     Temp(m) = Quote(String(P1, " "))
                        Case "str(":       Temp(m) = Quote("" & P1)
                        Case "substring(": Temp(m) = Quote(Mid(s1, P2, P3))
                        Case "upper(":     Temp(m) = Quote(UCase(s1))
                        
                        Case Else
                            Err.Raise EQ_FUNCTION, "clsEquation", "Undefined Function: " & v & ")"
                            Exit Sub
                    End Select
                Else
                    Select Case v
                        Case "pi":       Temp(m) = PI
                        Case "e":        Temp(m) = 2.718281828
                        Case "rnd":      Temp(m) = Rnd
                        Case Else
                            If Left(v, 1) = """" Then
                                Temp(m) = v
                            Else
                                eSpace = ER_VAR
                                Temp(m) = CDbl(Vars(Temp(m)))
                                eSpace = ER_NONE
                            End If
                    End Select
                End If
        End Select
    Next i
   
    dAnswer = CDbl(Temp(GetRight(0, Temp)))
    Dirty = False
    Exit Sub

SolveError:
    Select Case Err
        Case 6, 11, EQ_PAREN To EQ_NAME
            Err.Raise Err, "clsEquation", Err.Description
        Case 5:
            Select Case eSpace
                Case ER_VAR
                    Err.Raise EQ_VARIABLE, "clsEquation", "Undefined Function: " & v
                Case Else
                    Err.Raise Err, "clsEquation", Err.Description
            End Select
        Case Else
            Err.Raise EQ_INVALID, "clsEquation", "Invalid Equation"
    End Select
End Sub

Private Sub Parse()
    Dim i                   As Integer
    Dim s                   As String
    Dim t                   As Integer
    Dim j                   As Integer
    Dim sTmp                As String
    Dim p                   As Integer
    Dim EquPriority         As Collection
    Dim maxPriority         As Integer
    Dim isNeg               As Boolean
    Dim isQuote             As Boolean
    
    On Local Error GoTo Err_Parse
    
    s = ""
    t = EQ_NONE
    j = 1
    p = 0
    isNeg = False
    Set EquPriority = New Collection
    Set EquParsed = New Collection
    EquParsed.Add ""
    EquPriority.Add ""
    maxPriority = PRI_LEVEL
    
    ' added to support 'mod' in addition to %
    Equ = Replace(Equ, " mod ", "%")
    ' added to support <=,>=
    Equ = Replace(Equ, "<=", Chr(2))
    Equ = Replace(Equ, ">=", Chr(3))
    Equ = Replace(Equ, "<>", Chr(4))
    
    
    Equ = Replace(Equ, " AND ", Chr(5))
    Equ = Replace(Equ, " AND(", Chr(5) & "(")
    Equ = Replace(Equ, ")AND ", ")" & Chr(5))
    Equ = Replace(Equ, ")AND(", ")" & Chr(5) & "(")
    
    Equ = Replace(Equ, " OR ", Chr(6))
    Equ = Replace(Equ, " OR(", Chr(6) & "(")
    Equ = Replace(Equ, ")OR ", ")" & Chr(6))
    Equ = Replace(Equ, ")OR(", ")" & Chr(6) & "(")
    
    
    For i = 1 To Len(Equ)
        sTmp = Mid$(Equ, i, 1)
        Select Case sTmp
            Case "A" To "Z", "a" To "z", "_", """", IIf(isQuote, sTmp, "xxx")
                If sTmp = """" Then isQuote = Not isQuote
                If t = EQ_NONE Then
                    t = EQ_STRING
                    s = sTmp
                ElseIf t = EQ_NUMBER Then
                    t = EQ_STRING
                    EquParsed.Add s, , j
                    EquPriority.Add 0, , j
                    j = j + 1
                    EquParsed.Add "*", , j
                    EquPriority.Add PRI_MUL + p, , j
                    j = j + 1
                    s = sTmp
                Else
                    s = s + sTmp
                End If
                isNeg = True
        
              Case "1" To "9", "0", "."
                    If t = EQ_NONE Then
                        t = EQ_NUMBER
                        s = sTmp
                    Else
                        s = s + sTmp
                    End If
                    isNeg = True
                    
              Case "(":
                    If t = EQ_STRING Then
                        EquParsed.Add s + sTmp, , j
                        EquPriority.Add p + PRI_PAR, , j
                        j = j + 1
                        s = ""
                    ElseIf t = EQ_NUMBER Then
                        EquParsed.Add s, , j
                        EquPriority.Add 0, , j
                        j = j + 1
                        EquParsed.Add "*", , j
                        EquPriority.Add p + PRI_MUL, , j
                        j = j + 1
                        EquParsed.Add sTmp, , j
                        EquPriority.Add p + PRI_PAR, , j
                        j = j + 1
                        s = ""
                    Else
                        EquParsed.Add sTmp, , j
                        EquPriority.Add p + PRI_PAR, , j
                        j = j + 1
                    End If
                    p = p + PRI_LEVEL
                    t = EQ_NONE
                    If maxPriority < p + PRI_LEVEL Then
                        maxPriority = p + PRI_LEVEL
                    End If
                    isNeg = False
                    
                Case "=", "<", ">", Chr(2), Chr(3), Chr(4)  'operators
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_OPERATOR, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                
                Case Chr(5), Chr(6) ' booleans
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_BOOLEAN, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                
                Case "*", "/":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_MUL, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                    
                Case "\":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_MUL, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                    
                Case "+"
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                        EquParsed.Add sTmp, , j
                        EquPriority.Add p + PRI_ADD, , j
                        j = j + 1
                        t = EQ_NONE
                    Else
                        ' added by dd
                        EquParsed.Add sTmp, , j
                        EquPriority.Add p + PRI_ADD, , j
                        j = j + 1
                        t = EQ_NONE
                    End If
                    isNeg = False
                Case "-":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    If isNeg Then
                        EquParsed.Add sTmp, , j
                        EquPriority.Add p + PRI_ADD, , j
                        j = j + 1
                        t = EQ_NONE
                    Else
                        EquParsed.Add "~", , j
                        EquPriority.Add p + PRI_NEG, , j
                        j = j + 1
                        t = EQ_NONE
                    End If
                    isNeg = False
                Case "^":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_EXP, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                Case "%":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p + PRI_MOD, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                Case ",":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add Null, , j
                    EquPriority.Add 0, , j
                    j = j + 1
                    t = EQ_NONE
                    isNeg = False
                Case ")":
                    If t <> EQ_NONE Then
                        EquParsed.Add s, , j
                        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
                        j = j + 1
                        s = ""
                    End If
                    EquParsed.Add sTmp, , j
                    EquPriority.Add p - (PRI_LEVEL - PRI_PAR), , j
                    p = p - PRI_LEVEL
                    j = j + 1
                    t = EQ_NONE
                    isNeg = True
            End Select
    Next i
        
    If s <> "" Then
        EquParsed.Add s, , j
        EquPriority.Add IIf(t = EQ_STRING, p + PRI_VAR, 0), , j
        j = j + 1
    End If
    EquParsed.Remove j
    EquPriority.Remove j
    If p <> 0 Then
        Err.Raise EQ_PAREN, "clsEquation", "Unbalanced parenthesis"
        Set EquPriority = Nothing
        Exit Sub
    End If
    Set EquOrder = New Collection
    EquOrder.Add ""
    For j = 1 To maxPriority
        For i = EquPriority.Count To 1 Step -1
            If EquPriority(i) = j Then
                EquOrder.Add i, , , 1
            End If
        Next i
    Next j
    EquOrder.Remove 1
    Parsed = True
    Set EquPriority = Nothing
    
    Exit Sub

Err_Parse:
    Set EquPriority = Nothing
    Err.Raise Err, "clsEquation.Parse", Err.Description
    
End Sub

Private Function GetRight(ByVal j As Long, v() As Variant) As Long
    ' helper function for solve
    Dim i As Long
    For i = j + 1 To UBound(v)
        If Not IsNull(v(i)) Then
            GetRight = i
            Exit Function
        End If
    Next i
    GetRight = 0
End Function

Private Function GetLeft(ByVal j As Long, v() As Variant) As Long
    ' helper function for solve
    Dim i As Long
    
    For i = j - 1 To 1 Step -1
        If Not IsNull(v(i)) Then
            GetLeft = i
            Exit Function
        End If
    Next i
    GetLeft = 0
End Function




Private Function trunc(Number As Double, ByVal DecimalPlaces As Double) As Double
    DecimalPlaces = -1 * Int(DecimalPlaces)
    trunc = Int(Number / 10 ^ DecimalPlaces) * 10 ^ DecimalPlaces
End Function

Private Function Round(ByVal Number As Double, ByVal DecimalPlaces As Double) As Double



    DecimalPlaces = -1 * Int(DecimalPlaces)
    Round = VBA.Round(Int(Number / 10 ^ (DecimalPlaces - 1)) / 10) * 10 ^ DecimalPlaces
    
    'Round = Format(Number, "0." & String(Int(DecimalPlaces), "0"))
End Function


Private Function RoundTo(ByVal Number, Method, Significance) As Double
    
    
    Dim tmp As Integer
    Dim tmpVal As Double
    Dim TmpVal2 As Double
    
    If Significance = 0 Or Number = Significance Then
        RoundTo = Number
        Exit Function
    End If
    If Number / Significance = Int(Number / Significance) Then
        RoundTo = Number
        Exit Function
    End If
    Select Case Method

        Case 1 ' up
            tmpVal = ((Number / Significance) + (-0.5 + IIf(Method = 1, 1, 0)))
            TmpVal2 = tmpVal
            tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)
            Number = Fix(TmpVal2) + tmpVal / 10 ^ 0
            RoundTo = Number * Significance

        Case 0 ' closest
            tmpVal = RoundTo(Number, -1, Significance)
            TmpVal2 = RoundTo(Number, 1, Significance)
            RoundTo = IIf(Abs(Number - tmpVal) < Abs(Number - TmpVal2), tmpVal, TmpVal2)
        
        Case -1 ' down
            'If Significance < 1 Then
            '    RoundTo = Int(Number) + Significance
            'Else
                tmpVal = ((Number / Significance) + (-0.5 + IIf(Method = 1, 1, 0)))
                TmpVal2 = tmpVal
                tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)
                Number = Fix(TmpVal2) + tmpVal / 10 ^ 0
                RoundTo = Number * Significance
            'End If
            

    End Select

    

End Function

Private Function Min(ParamArray items())
    Dim i As Long
    Dim v As Variant
    v = items(LBound(items))
    For i = LBound(items) + 1 To UBound(items)
        v = IIf(items(i) > v, v, items(i))
    Next
    Min = v
End Function
Private Function Max(ParamArray items())
    Dim i As Long
    Dim v As Variant
    v = items(LBound(items))
    For i = LBound(items) + 1 To UBound(items)
        v = IIf(items(i) > v, items(i), v)
    Next
    Max = v
 End Function

Private Function sum(ParamArray items())
    Dim i As Long
    Dim v As Variant
    v = 0
    For i = LBound(items) To UBound(items)
        v = v + items(i)
    Next
    sum = v
End Function


Private Function Avg(ParamArray items())
    Dim i As Long
    Dim X As Long
    Dim v As Variant
    v = 0
    X = 0
    For i = LBound(items) To UBound(items)
        X = X + 1
        v = v + items(i)
    Next
    Avg = v / X
End Function







Private Function UnQuote(s) As String
    If Left(s, 1) = vbQuote Then s = Mid(s, 2)
    If Right(s, 1) = vbQuote Then s = Left(s, Len(s) - 1)
    UnQuote = s
End Function
Private Function Quote(s) As String
    Quote = vbQuote & s & vbQuote
End Function


Private Function temptables(Temp()) As String
    Dim i As Long
    Dim s As String
    
    s = "EquParsed" & vbCrLf
    For i = 1 To UBound(Temp)
        s = s & " " & i & ".  " & Temp(i) & vbCrLf
    Next
    
    s = s & vbCrLf & vbCrLf & "EquOrder" & vbCrLf
    For i = 1 To EquOrder.Count
        s = s & " " & EquOrder(i) & ".  " & EquParsed(EquOrder(i)) & vbCrLf
    Next
    
    temptables = s
    
End Function




Private Function Ceiling(ByVal Number) As Long
    If Number >= 0 Then
        If Number = Int(Number) Then
            Ceiling = Number
        Else
            Ceiling = Int(Number) + 1
        End If
    ElseIf Number < 0 Then
        Ceiling = Int(Number)
    End If
End Function


Private Function Floor(ByVal Number) As Long
    Floor = Fix(Number)
    If Number >= 0 Then
        If Number = Int(Number) Then
            Floor = Number
        Else
            Floor = Int(Number)
        End If
    ElseIf Number < 0 Then
        Floor = Int(Number) - 1
    End If
End Function

