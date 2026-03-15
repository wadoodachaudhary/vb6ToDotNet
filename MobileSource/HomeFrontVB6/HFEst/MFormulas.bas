Attribute VB_Name = "MFormulas"
Option Explicit

Public Function ResolveFormula(ByVal RawFormula As String) As String
    Dim originalformula As String
    Dim rs As Recordset
    Dim Name As String
    
    Dim fnames As String
    Dim i As Long
    
    'replace {formulaname} with code
    originalformula = RawFormula
    
    
    RawFormula = Replace(RawFormula, vbCr, " ")
    RawFormula = Replace(RawFormula, vbLf, " ")
    RawFormula = Trim(RawFormula)
    While InStr(1, RawFormula, "{", vbTextCompare)
        
        Name = Parse(Parse(RawFormula, 2, "{"), 1, "}")
        
        'check for circular ref
        i = i + 1
        fnames = fnames & Name & vbCrLf
        If i > 30 Then
            MsgBox "Formula cannot be resolved." & vbCrLf & vbCrLf & originalformula & vbCrLf & vbCrLf & "Call Stack:" & vbCrLf & fnames, vbCritical, App.ProductName
            ResolveFormula = ""
            Exit Function
        End If
        
        
        Set rs = HFApp.SqlExec("Select text from Formulas where name = " & DbQuote(Str, Name))
        If rs.EOF Then
            MsgBox "Unknown formula " & Name, vbExclamation, App.ProductName
            Exit Function
        Else
            RawFormula = Replace(RawFormula, "{" & Name & "}", "(" & rs(0) & ")")
        End If
    Wend
    
    
    ResolveFormula = RawFormula
End Function


Private Function PrefixFunction(Formula As String, FunctionName As String) As String
    Dim re As RegExp
    Dim Matches As MatchCollection
    Dim Match As Match
    Dim m As Long
    
    'replaces FunctionName with dbo.FunctionName
    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    
    
    
    'first remove any "dbo."'s the user may have typed
    Formula = Replace(Formula, "dbo." & FunctionName, FunctionName, , , vbTextCompare)
    
    
    'now add dbo.
    re.Pattern = "\b" & FunctionName & "\b"
    Set Matches = re.Execute(Formula)
    For m = Matches.Count To 1 Step -1
        Set Match = Matches.Item(m - 1)
        If Trim(Match.value) <> "" Then
            Formula = Left(Formula, Match.FirstIndex) & "dbo." & FunctionName & Mid(Formula, Match.FirstIndex + Match.Length + 1)
        End If
    Next

    PrefixFunction = Formula

End Function


Public Function ParseVariables(ByVal ResolvedFormula As String) As String
    Dim s As String
    Dim list As String
    Dim Name As String
    
    While InStr(1, ResolvedFormula, "[", vbTextCompare)
        Name = Parse(Parse(ResolvedFormula, 2, "["), 1, "]")
        list = list & Chr(1) & Name
        
        s = ResolvedFormula
        ResolvedFormula = Replace(ResolvedFormula, "[" & Name & "]", "", , , vbTextCompare)
        If s = ResolvedFormula Then ResolvedFormula = ""
        
    Wend
    
    ParseVariables = Mid(list, 2)
End Function

Public Sub LoadVariables(Clear As Boolean, Grid As VSFlexGrid, variablelist As String)
    
    'VariableList is a chr(1) delimited list of names without []
    Dim i As Long
    Dim r As Long
    Dim Name As String
    Dim rs As Recordset
    Dim s As String
    
    
    With Grid
        
        If Clear Then .Rows = 0
        
        For i = 0 To Parse(variablelist, , Chr(1)) - 1
            Name = Parse(variablelist, i + 1, Chr(1))
            If Name <> "" Then
            If .FindRow(Name, 0, 0, False) = -1 Then
            
            
                Set rs = HFApp.SqlExec("select * from variables where name=" & DbQuote(Str, Name), dbHomefront)
                .AddItem ""
                r = .Rows - 1
                .TextMatrix(r, .ColIndex("Name")) = Name
                If rs.EOF Then
                    .TextMatrix(r, .ColIndex("Min")) = -1E+18
                    .TextMatrix(r, .ColIndex("Max")) = 1E+18
                Else
                    s = " |" & Trim("" & rs("listofvalues"))
                    If s = " |" Then
                        .TextMatrix(r, .ColIndex("Value")) = Val("" & rs("defaultvalue"))
                        .TextMatrix(r, .ColIndex("UOM")) = "" & rs("uom")
                        .Cell(flexcpData, r, .ColIndex("UOM")) = "" & rs("uom")  'this is used on the takeoff window. stuff is editable if the "original" uom is empty
                        .TextMatrix(r, .ColIndex("Min")) = Val("" & rs("minimumvalue"))
                        .TextMatrix(r, .ColIndex("Max")) = Val("" & rs("maximumvalue"))
                    Else
                        .TextMatrix(r, .ColIndex("Value")) = Parse(s, 2, "|")
                        .TextMatrix(r, .ColIndex("ListOfValues")) = s
                    End If
                    .TextMatrix(r, .ColIndex("Help")) = "" & rs("help")
                    .TextMatrix(r, .ColIndex("ConditionName")) = "" & rs("ConditionName")
                    .TextMatrix(r, .ColIndex("ConditionType")) = "" & rs("ConditionType")
                    .TextMatrix(r, .ColIndex("ConditionUOM")) = "" & rs("ConditionUOM")
                End If
            End If
            End If
        Next
        On Error Resume Next
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = vbButtonFace
        .Cell(flexcpBackColor, 0, 2, .Rows - 1, 2) = vbButtonFace
    End With
End Sub

Public Function OLDCalcFormula(Grid As VSFlexGrid, ByVal ResolvedFormula As String, Quiet As Boolean) As Double
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    s = ResolvedFormula
    With Grid
        For i = 0 To .Rows - 1
            If .TextMatrix(i, .ColIndex("ListOfValues")) <> "" Then
                s = Replace(s, "[" & .TextMatrix(i, .ColIndex("name")) & "]", DbQuote(Str, .TextMatrix(i, .ColIndex("value"))), , , vbTextCompare)
            Else
                s = Replace(s, "[" & .TextMatrix(i, .ColIndex("name")) & "]", "cast(" & DbQuote(Num, .TextMatrix(i, .ColIndex("value"))) & " as float)", , , vbTextCompare)
            End If
        Next
    End With
    
    If s = "" Then
        OLDCalcFormula = 0
        Exit Function
    Else
        
        s = PrefixFunction(s, "RNDTO")
        s = PrefixFunction(s, "MIN")
        s = PrefixFunction(s, "MAX")
        s = PrefixFunction(s, "AVG")
        s = PrefixFunction(s, "TRUNC")
        s = PrefixFunction(s, "MOD")
        
        s = "Select " & s
        s = HFApp.SqlExec(s)(0)
        OLDCalcFormula = s
    End If
    
Exit Function
eh:
    If Not Quiet Then
        Select Case True
            Case Err.Number = 13 'type mismatch
                MsgBox "Formula returns '" & s & "'. It must return a number.", vbExclamation, "Error Evaluating Formula"
            Case InStr(1, Err.Description, "Divide by zero", vbTextCompare)
                MsgBox "Divide by zero error encountered.", vbExclamation, "Error Evaluating Formula"
            Case Else
                MsgBox Parse(Err.Description, Parse(Err.Description, , "[SQL Server]"), "[SQL Server]") & vbCrLf & vbCrLf & s, vbExclamation, "Error Evaluating Formula"
        End Select
    End If
    On Error GoTo 0
    Err.Raise 9
End Function

Public Function FormulaHasLookups(ResolvedFormula As String) As Boolean
    FormulaHasLookups = InStr(1, ResolvedFormula, "TotalTakeoffQty", vbBinaryCompare) Or InStr(1, ResolvedFormula, "TotalOrderQty", vbBinaryCompare) Or InStr(1, ResolvedFormula, "TotalCost", vbBinaryCompare)
End Function

Public Function CalcFormula(variables As VSFlexGrid, ByVal ResolvedFormula As String, Quiet As Boolean, ParentForm As Object) As Double
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim X As New clsEquation
Dim value As Double
    
    'replace [VariableName] with value
    s = ResolvedFormula
    With variables
        For i = 0 To .Rows - 1
            If .TextMatrix(i, .ColIndex("ListOfValues")) <> "" Then
                s = Replace(s, "[" & .TextMatrix(i, .ColIndex("name")) & "]", """" & .TextMatrix(i, .ColIndex("value")) & """", , , vbTextCompare)
            Else
                s = Replace(s, "[" & .TextMatrix(i, .ColIndex("name")) & "]", Val(.TextMatrix(i, .ColIndex("value"))), , , vbTextCompare)
            End If
        Next
    End With
    
    
    
    If s = "" Then
        CalcFormula = 0
        Exit Function
    Else
        
        X.Equation = s
        Set X.ParentForm = ParentForm
        value = X.Solution
        CalcFormula = value
    End If
    
Exit Function
eh:
    If Not Quiet Then
        Select Case True
            Case Err.Number = 13 'type mismatch
                MsgBox "Formula returns '" & s & "'. It must return a number.", vbExclamation, "Error Evaluating Formula"
            Case InStr(1, Err.Description, "Divide by zero", vbTextCompare)
                MsgBox "Divide by zero error encountered.", vbExclamation, "Error Evaluating Formula"
            Case Else
                MsgBox Parse(Err.Description, Parse(Err.Description, , "[SQL Server]"), "[SQL Server]") & vbCrLf & vbCrLf & s, vbExclamation, "Error Evaluating Formula"
        End Select
    End If
    On Error GoTo 0
    Err.Raise 9
End Function

