Attribute VB_Name = "IntelliSenseClone"
Option Explicit

Public Sub IntelliSense(txtbox As Object, KeyAscii As Integer, Hinimain As Hini, ListCompatable As Object)
    Dim temp As String, tempstr As String, templ As Long
    Select Case KeyAscii
        Case 8, 13, 46 'GNDN
        Case Asc("("), 57 'Check for function
        Case Asc("."), 190 'Check for variable
            temp = CurrWord(txtbox, txtbox.SelStart - 1)
            
            tempstr = txtbox.text
            templ = txtbox.SelStart
            
            ListCompatable.clear
            
            If txtbox <> tempstr Then
                txtbox = tempstr
                txtbox.SelStart = templ
            End If
            
            If VarExists(temp, Hinimain) Then
                temp = getType(temp, Hinimain)
                SeedList ListCompatable, "Types\" & temp, Hinimain
                ListCompatable.Refresh
                ListCompatable.Visible = True
            End If

        Case Else
            AutoComplete txtbox, ListCompatable
    End Select
End Sub
Public Sub AutoComplete(txtbox As Object, ListCompatable As Object)
    Dim temp As String, count As Long, found As String
    temp = CurrWord(txtbox)
    For count = 0 To ListCompatable.ListCount - 1
        If StrComp(Left(ListCompatable.List(count), Len(temp)), temp, vbTextCompare) = 0 Then
            If Len(found) = 0 Then
                found = ListCompatable.List(count)
                'ListCompatable.ListIndex = count
            Else
                Exit Sub
            End If
        End If
    Next
    If Len(found) > 0 Then
        found = Right(found, Len(found) - Len(temp))
        count = txtbox.SelStart
        temp = Left(txtbox, count) & found & Right(txtbox, Len(txtbox) - count)
        txtbox.text = temp
        txtbox.SelStart = count
        txtbox.SelLength = Len(found)
        If Not ListCompatable.Visible Then ListCompatable.Visible = True
    End If
End Sub
Public Sub SeedList(ListCompatable As Object, Section As String, Hinimain As Hini, Optional doSubs As Boolean)
    Dim temp As Long, tempstr() As String, count As Long
    count = Hinimain.keycount(Section)
    Hinimain.enumeratekeys Section, tempstr
    For temp = 1 To count
        ListCompatable.additem tempstr(1, temp)
    Next
    If doSubs Then
        Dim tempstr2() As String
        Hinimain.enumeratesections Section, tempstr2
        For temp = 1 To Hinimain.sectioncount(Section)
            ListCompatable.additem tempstr2(temp)
        Next
    End If
End Sub
Public Function CurrWord(txtbox As Object, Optional startfrom As Long = -1) As String
    Dim start As Long, finish As Long
    With txtbox
        If startfrom = -1 Then startfrom = .SelStart
        start = getNextDel(.text, startfrom, -1) + 1
        finish = getNextDel(.text, startfrom)
        If finish > start Then CurrWord = Mid(txtbox, start, finish - start)
    End With
End Function
Public Function ForceAutoComplete(txtbox, newword As String, Optional startfrom As Long = -1) As String
    Dim start As Long, finish As Long, temp As String
    With txtbox
        If startfrom = -1 Then startfrom = .SelStart
        start = getNextDel(.text, startfrom, -1)
        finish = getNextDel(.text, startfrom + 1)
        If start = 0 And finish = Len(txtbox) + 1 Then
            ForceAutoComplete = newword
        Else
        
        If finish >= start Then
            temp = Left(.text, start) & newword
            If finish < Len(txtbox) Then temp = temp & Right(txtbox, Len(txtbox) - finish)
            ForceAutoComplete = temp
            'ForceAutoComplete = Left(.text, start) & newword & Right(.text, Len(.text) - finish)
        End If
        
        End If
    End With
End Function
Public Function getNextDel(text As String, Optional ByVal start As Long = 1, Optional direction As Long = 1) As Long
    On Error Resume Next
    Do Until IsADelimeter(Mid(text, start, 1))
        If start < 1 Then
            getNextDel = 0
            Exit Function
        End If
        If start > Len(text) Then
            getNextDel = Len(text) + 1
            Exit Function
        End If
        start = start + direction
    Loop
    getNextDel = start
End Function
Private Function IsADelimeter(char As String) As Boolean
    Select Case char
        Case vbNewLine, vbTab, " ", ",", ".": IsADelimeter = True
    End Select
End Function
