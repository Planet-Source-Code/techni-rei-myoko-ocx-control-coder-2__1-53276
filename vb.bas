Attribute VB_Name = "VisualBasicFileFunctions"
Option Explicit
Private Const builtintypes As String = "boolean,byte,date,double,integer,long,object,single,string,variant,ole_cancelbool,ole_color,ole_handle,ole_optexclusive"
Private Const properties As Long = 0, code As Long = 1
Private currsub As String
Public Sub ScanFile(filename As String, Hinimain As Hini)
    Dim tempfile As Long, linenumber As Long, tempstr As String, temparr() As String, mode As Long
    Dim currcontrol As String
    With Hinimain
        currsub = Empty
        .creationsection "Variables"
        .creationsection "Functions"
        AddBuiltInTypes Hinimain
        If FileLen(filename) > 0 Then
            tempfile = FreeFile
            Open filename For Input As #tempfile
                Do Until EOF(tempfile)
                    linenumber = linenumber + 1
                    tempstr = LCase(LInput(tempfile))
                    If Len(tempstr) > 0 Then
                    temparr = Split(tempstr, " ")
                    If mode = properties Then
                        Select Case temparr(0)
                            Case "version"
                                If temparr(1) <> "5.00" And linenumber = 1 Then
                                    MsgBox "Incorrect version number", vbCritical, "An error occurred"
                                    Exit Sub
                                End If
                            Case "begin"
                                AddType tempfile, temparr(1), Hinimain
                                AddVariable temparr(2), temparr(1), Hinimain
                                mode = code
                        End Select
                    Else
                        Select Case temparr(0)
                            Case "public", "private", "dim" 'add property, sub, function, type,enum,api declare,const,var
                                Select Case temparr(1)
                                    Case "sub", "function", "property", "declare", "event"
                                        AddSub temparr, Hinimain
                                    Case "type", "enum"
                                        AddType tempfile, temparr(2), Hinimain
                                    Case Else
                                        AddVars temparr, Hinimain
                                End Select
                        End Select
                    End If
                    End If
                    DoEvents
                Loop
            Close #tempfile
        End If
    End With
End Sub
Public Sub AddBuiltInTypes(Hinimain As Hini)
    Dim tempstr() As String, temp As Long
    tempstr = Split(builtintypes, ",")
    Hinimain.creationsection "Types"
    For temp = 0 To UBound(tempstr)
        Hinimain.creationsection "Types\" & tempstr(temp)
    Next
End Sub
Public Function CleanBracket(text As String) As Boolean
    If Left(text, 1) = "(" Then
        CleanBracket = True
        text = Right(text, Len(text) - 1)
    End If
    If Right(text, 1) = ")" Then
        CleanBracket = True
        text = Left(text, Len(text) - 1)
        If Right(text, 1) = "(" Then
            CleanBracket = False
            text = Left(text, Len(text) - 1)
        End If
    End If
End Function
Public Function SubExists(routine As String, Hinimain As Hini)
    SubExists = Hinimain.existancesection("Functions\" & routine)
End Function
Public Sub AddSub(temparray, Hinimain As Hini)
    Dim temp As Long, Name As String, mytype As String, isend As Boolean
    For temp = 0 To UBound(temparray)
        Select Case temparray(temp)
            Case "public", "private", "property" 'declare scope
            Case "declare"
                Name = temparray(temp + 2)
                temp = temp + 4
            Case "event"
                temp = temp + 1
                Name = Left(temparray(temp), InStr(temparray(temp), "(") - 1)
                AddEvent Name, temparray, temp, Hinimain
            Case "sub", "function", "let", "get", "set"
                temp = temp + 1
                Name = Left(temparray(temp), InStr(temparray(temp), "(") - 1)
                currsub = Name
                'temparray(temp) = Right(temparray(temp), Len(temparray(temp)) - InStr(temparray(temp), "("))
                AddParameter2Sub Name, temparray, temp, Hinimain
         End Select
    Next
End Sub
Public Sub AddEvent(Name As String, temparray, start As Long, Hinimain As Hini)
'GNDN
End Sub
Public Sub AddParameter2Sub(routine As String, temparray, start As Long, Hinimain As Hini)
On Error Resume Next
Dim alert As Long
Dim Name As String, endit As Boolean, temp As String, temp2 As String
Hinimain.creationsection "Functions\" & routine
Do Until start > UBound(temparray) Or endit
    '(temp as long, tempw, paramarray temparr() as variant)as quax
    alert = InStr(temparray(start), "(")
    If alert > 0 And alert < Len(temparray(start)) Then
        If InStr(1, temparray(start), "byval", vbTextCompare) > 0 Or _
            InStr(1, temparray(start), "byref", vbTextCompare) > 0 Or _
            InStr(1, temparray(start), "optional", vbTextCompare) > 0 Or _
            InStr(1, temparray(start), "paramarray", vbTextCompare) > 0 Then
                temparray(start) = Right(temparray(start), Len(temparray(start)) - alert)
        End If
    End If
    Select Case temparray(start)
        Case "optional", "byref", "byval" 'GNDN
        Case "paramarray"
            AddParameter routine, CStr(temparray(start + 1)), "Variant", Hinimain
            If temparray(start + 2) = "as" Then
                start = start + 4
            End If
        Case "as"
            AddParameter routine, routine, CStr(temparray(start + 1)), Hinimain
        Case Else
            If temparray(start) = "optional" Then start = start + 1
            If temparray(start) = "byref" Then start = start + 1
            If temparray(start) = "byval" Then start = start + 1
            
            If temparray(start + 1) = "as" Then
                temp = temparray(start + 2)
                AddParameter routine, CStr(temparray(start)), temp, Hinimain
                start = start + 2
            Else
                AddParameter routine, CStr(temparray(start)), "Variant", Hinimain
            End If
    End Select
    start = start + 1
    DoEvents
Loop
End Sub
Public Sub AddParameter(routine As String, Parameter As String, pType As String, Hinimain As Hini)
    Dim temp As Long
    temp = InStr(Parameter, "()")
    If temp > 0 Then Parameter = Left(Parameter, temp - 1)
    temp = InStr(Parameter, "(")
    If temp > 0 Then Parameter = Right(Parameter, Len(Parameter) - temp)
    temp = InStr(pType, ")")
    If temp = Len(pType) Then pType = Left(pType, temp - 1)
    Parameter = Replace(Parameter, ",", Empty)
    pType = Replace(pType, ",", Empty)
    Hinimain.setkeycontents "Functions\" & routine, Parameter, pType
End Sub
Public Sub AddVars(ByVal temparray, Hinimain As Hini)
    Dim temp As Long, tempstr As String
    Dim Vname As String, vType As String
    temparray(0) = Empty 'Should be dim, private or public
    tempstr = Join(temparray, " ")
    tempstr = Right(tempstr, Len(tempstr) - 1)
    temparray = Split(tempstr, ", ")
    For temp = 0 To UBound(temparray)
        vType = "variant"
        Vname = temparray(temp)
        If InStr(Vname, " as ") > 0 Then
            vType = Right(Vname, Len(Vname) - InStrRev(Vname, " "))
            Vname = Left(Vname, InStr(Vname, " ") - 1)
        End If
        If Len(currsub) = 0 Then
            AddVariable Vname, vType, Hinimain
        Else
            If Not Hinimain.existancesection("Variables\" & currsub) Then
                Hinimain.creationsection "Variables\" & currsub
            End If
            Hinimain.setkeycontents "Variables\" & currsub, Vname, vType
        End If
    Next
End Sub
Public Sub AddVariable(Name As String, vType As String, Hinimain As Hini)
    If InStr(Name, " ") = 0 Then Hinimain.setkeycontents "Variables", Name, vType
End Sub
Public Function VarExists(Name As String, Hinimain As Hini, Optional routine As String) As Boolean
    If Len(routine) = 0 Then
        VarExists = Hinimain.existsancekey("Variables", Name)
    Else
        VarExists = Hinimain.existsancekey("Variables\" & routine, Name)
    End If
End Function
Public Function LInput(Filenumber As Long) As String
    Dim tempstr As String
    If EOF(Filenumber) Then Exit Function
    Line Input #Filenumber, tempstr
    tempstr = Trim(Replace(tempstr, vbTab, Empty))
    Do Until InStr(tempstr, "  ") = 0
        tempstr = Replace(tempstr, "  ", " ")
        DoEvents
    Loop
    If Right(tempstr, 1) = "_" Then tempstr = Left(tempstr, Len(tempstr) - 1) & LInput(Filenumber)
    LInput = tempstr
End Function
Private Sub AddProperty(TypeName As String, PropName As String, Hinimain As Hini)
    PropName = Replace(PropName, "=", Empty)
    Hinimain.setkeycontents "Types\" & TypeName, PropName, Empty
End Sub
Public Function TypeExists(Name As String, Hinimain As Hini)
    TypeExists = Hinimain.existancesection("Types\" & Name)
End Function
Public Function PropertyCount(vType As String, Hinimain As Hini) As Long
    PropertyCount = Hinimain.keycount("Types\" & vType)
End Function
Public Function getType(Vname As String, Hinimain As Hini, Optional routine As String) As String
    If Len(routine) = 0 Then
        getType = Hinimain.getkeycontents("Variables", Vname)
    Else
        getType = Hinimain.getkeycontents("Variables\" & routine, Vname)
    End If
End Function
Private Sub AddType(Filenumber As Long, Name As String, Hinimain As Hini)
    Dim tempstr As String, finished As Boolean, temparr() As String
    If Not TypeExists(Name, Hinimain) Then
        Hinimain.creationsection "Types\" & Name
        Do Until finished Or EOF(Filenumber)
            tempstr = LCase(LInput(Filenumber))
            temparr = Split(tempstr, " ")
            Select Case temparr(0)
                Case "begin"
                    AddType Filenumber, temparr(1), Hinimain
                    AddVariable temparr(2), temparr(1), Hinimain
                Case "beginproperty"
                    AddType Filenumber, temparr(1), Hinimain
                    AddProperty Name, temparr(1), Hinimain
                Case "end", "endproperty"
                    finished = True
                Case Else
                    AddProperty Name, temparr(0), Hinimain
            End Select
            DoEvents
        Loop
    Else
        Do Until finished Or EOF(Filenumber)
            tempstr = LCase(LInput(Filenumber))
            If tempstr = "end" Or tempstr = "endproperty" Then
                finished = True
            End If
        Loop
    End If
End Sub
Private Function stripcode(code As String) As String
    Dim temp As Long, tempstr As String
    temp = InStr(code, vbNewLine)
    tempstr = Right(code, Len(code) - temp)
    temp = InStrRev(tempstr, vbNewLine)
    tempstr = Left(tempstr, temp)
    stripcode = tempstr
End Function
Public Function SetFunction(ByVal file As String, routine As String, ByVal code As String, Optional clear As Boolean) As String
    Dim start As Long, finish As Long
    start = StartOfFunction(file, routine)
    If start > 0 Then
        code = stripcode(code)
        start = InStr(start, file, vbNewLine) + 2
        finish = EndOfFunction(file, start) - 1
        If Not clear Then code = Trim(Mid(file, start, finish - start)) & code
        file = Left(file, start) & code & Right(file, Len(file) - finish)
    Else
        file = file & vbNewLine & vbNewLine & code
    End If
    SetFunction = file
End Function
Public Function StartOfFunction(file As String, routine As String) As Long
    Dim temp As Long
    temp = InStr(1, file, "sub " & routine, vbTextCompare)
    If temp > 0 Then StartOfFunction = temp: Exit Function
    temp = InStr(1, file, "function " & routine, vbTextCompare)
    If temp > 0 Then StartOfFunction = temp: Exit Function
    temp = InStr(1, file, "property " & routine, vbTextCompare)
    StartOfFunction = temp
End Function
Public Function EndOfFunction(file As String, start As Long) As Long
    Dim temp(0 To 2) As Long, temp2 As Long, temp3 As Long
    temp(0) = InStr(start, file, "end sub", vbTextCompare)
    temp(1) = InStr(start, file, "end function", vbTextCompare)
    temp(2) = InStr(start, file, "end property", vbTextCompare)
    
    temp2 = temp(0)
    For temp3 = 1 To 2
        If temp(temp3) > 0 And temp(temp3) < temp2 Then temp2 = temp(temp3)
    Next
    EndOfFunction = temp2
End Function
