VERSION 5.00
Begin VB.UserControl Hini 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Hini2.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "Hini2.ctx":0C42
End
Attribute VB_Name = "Hini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Hierarchical INI format reading and editing functions
'Backwards compatable with standard INI files

'The Hini file format is similar to ini files as it contains groups of key=value
'pairs within [sections]. [section]s in hini files can now contain more [section]s
'which in turn can contain more key=value pairs and more [section]s and so forth
'a section is ended with a [/], known as a root, everything after that in no longer in the
'current [section]. Unlike ini files, the root itself can contain key=value pairs
'if a comment contains a '=' then it will be treated as a key, but this shouldnt matter much
'Can now have multiple keys using the same name as it can now check by instance/index

'Any time index is a parameter, and optional with a default of 0, its referring to the multikey instance
'if given 0 or 1, it'll use the first key with that name it encounters, and 2 is second, and so on

'Function name          Parameters                  Description

'The following are file manipulation commands
'savefile               filename                    Saves the loaded/created hini file to filename
'loadfile               filename                    Loads a hierarchical ini file
'closefile                                          purges the currently loaded hini file
'loadoldfile            filename                    Load an old/standard ini file as a hinifile, wont be converted to a hierarchy or multikey, [/] section will become root
'saveoldfile            filename                    Saves a hini file as an old/standard ini file, root keys will be put in the [/] section, multikeys will have their instance added in brackets
'loadxmlfile            filename                    Loads an XML file as a hini file
                                                    
'The following functions are the userfriendly versions of commands listed below

'setkeycontents         Section, Key, Value         Set the value of the key in the section to value
'getkeycontents         Section, Key, [Default]     Gets the contents of the key in the section, returns [default] if it doesnt exists
'keycount               Section                     Gets the number of keys in Section
'sectioncount           Section                     Gets the number of sections in Section
'keyname                Section, Key, Name          Sets the name of key in section to Name
'sectionname            Section, Name               Sets the name of Section to Name
'enumeratesections      Section, Array              Fills the array with: a list of sections within the section specified
'enumeratekeys          Section, Array              Fills the array with: a list of key names, a list of key values in the section specified
'enummultikey           Section, Key, Array         Fills the array with: a list of the contents of each key named key within the section specified
'existsancekey          Section, Key                Returns whether or not the key in the section exists
'existancesection       Section                     Returns whether or not the section exists
'creationsection        Section                     Creates every level of section in the section path if they dont exist
'createkey              Section, Key, Value         Creates a key in section with a value of value
'deletesection          Section                     Deletes section and all its contents
'deletekey              Section, Key                Deletes the key within the section

'keysindex              Section, Key                Returns the keys index number in the section
'sectionatindex         Section, Index              Returns the section with at index inside the section specified

'no command below this list is needed by the user as theyve been supplimented by ones above

'the following functions dont need a handle, but have been combined into the user friendly versions above
'newrootsection         section                     creates a new root section
'countrootsections                                  counts the sections not within other sections
'countrootkeys                                      counts keys in the root
'enumrootsections       array                       returns an array of root sections
'enumrootkeys           array                       returns an array of the keys in the root
'qualifiedsectionhandle '\' delimited section path  returns the handle to a fully qualified section path

'The following are involved with the new multikey features

'countinstancekeys      Start, Key                  Counts the number of times keys have the name of Key in the section at start
'counthandlekeys        Start, Index                Gets the handle of the key in the section at start at index Index
'getkeysinstance        Start, Handle               Gets the instance of the key in the section at start, at handle

'the following functions are used by the system to read hini files
'isroot                 String                      Returns if the text is a root [/] or not
'issection              String                      Returns if the text is a section [--text--] or not
'isnotroot              String                      Returns if the text is a section but not a root
'isvalue                String                      Returns if the text is a key=value pair
'iscomment              String                      Returns if the text is none of the above
'stripname              Key=value pair              Returns the Key portion
'stripvalue             Key=value pair              Returns the value portion
'stripsection           [section]                   Returns the section portion without the brackets

'the following functins have not been altered to use the qualifiedsectionhandle instead of start
'findroot               start                       searches for the next root
'countkeys              start                       counts the keys inside the section at start
'countsections          start                       counts the sections inside the section at start
'countmultikey          start, key                  counts the number of keys named key in the section at start
'handlerootsections     section                     returns the handle of the root section
'enumsections           start, array                returns an array of sections inside the section at start
'enumkeys               start  array                returns a multidimensional array (1 to 2, 1 to keycount) with the first array holding keynames, and the second containing the values
'sectionexists          start, section              returns if the section exists within the section at start
'sectionindex           start, section              returns the index number of the section within the section at start
'keyexists              start, key                  returns if the key exists within the section at start
'keyindex               start, index                returns the key at the index number in the section at start
'keyhandle              start, key                  returns the handle to the key in the section at start
'getkey                 start, key                  returns the value of key in the section at start
'setkey                 start, key, value           sets the value of key in the section at start to value
'sectionhandle          start, section              returns the handle to the section within the section at start
'renamekey              start, key, name            sets the name of the key in the section at start to name
'removesection          start                       removes the section at start
'renamesection          start, name                 renames the section at start to name
'createsection          start, section              creates a section in the section at start named section

'the following are system functions and recommended you not use
'insert                 start, level, contents      inserts a line at start with a level of level and the contents of contents
'removerange            top, bottom                 remove lines from top to bottom

'the following are debugging function that were used in the testing phase
'line                   start                       returns the level & ") " & contents of the line at start
'linecount                                          returns the number of lines

'These functions are used to convert hini files to ini files
'saveoldroot            filenumber                  Saves the root keys and sections as an ini file
'saveoldsection         filenumber                  Saves sub sections as root sections, removing hierarchy and multikey

'These functions are used to convert xml files to hini files
'loadwholefile          filename                    Returns the contents of a file
'findnext               text, start, char           Starting from start, it finds the next instance of char in text
'findprev               text, start, char           same as findnext but backwards
'removedoubles                                      Remove double section names, roots left unhandled as xml only allows on root (lucky me)
'removedoublessection   start                       Remove double section names (by adding index number) within the section at start

'A handle was used in most functions because this saves cpu time instead of getting it each time
'as many commands using the same handle are used in succession
'also I didnt make the qualifiedsectionhandle function till most of the functions were complete
'This code is by Techni Myoko, and anyone who claims otherwise is a liar
'this is directed at a specific user who enjoys stealing my code

Private Type entry
    level As Long
    contents As String
End Type
Private Enum errcode
    err_none = 0
    err_filenotloaded = 1
    err_filedoesntexist = 2
    err_sectionexists = 3
    err_sectiondoesntexist = 4
    err_keyexists = 5
    err_keydoesntexist = 6
    err_sectionhasnosections = 7
    err_sectionhasnokeys = 8
    err_filenotsaved = 9
End Enum

Dim stag As String, inifile() As entry, entrycount As Long, errorcode As errcode, isloaded As Boolean
Public Function loadwholefile(filename As String) As String
On Error Resume Next
If FileLen(filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(filename) <> filename Then
        Open filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                If tempstr2 <> Empty Then tempstr2 = tempstr2 & vbNewLine
                tempstr2 = tempstr2 & tempstr
                DoEvents
            Loop
            loadwholefile = tempstr2
        Close temp
    End If
End Function
Public Sub setkeycontents(Section As String, key As String, Value As String, Optional Index As Long = 0)
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        If keyexists(temp, key) = True Then
            setkey temp, key, Value, Index
        Else
            createkey Section, key, Value, Index
        End If
    End If
End Sub
Public Function getkeycontents(Section As String, key As String, Optional default As String, Optional Index As Long = 0) As String
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        'If keyexists(temp, key) Then
            getkeycontents = getkey(temp, key, Index)
        'Else
        '    getkeycontents = default
        'End If
    Else
        getkeycontents = default
    End If
End Function
Public Function keycount(Section As String) As Long
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        keycount = countkeys(temp)
    Else
        keycount = countrootkeys
    End If
End Function
Public Function sectioncount(Optional Section As String)
    If entrycount = 0 Then Exit Function
    If IsMissing(Section) Then
        sectioncount = countrootsections
    Else
        Dim temp As Long
        temp = qualifiedsectionhandle(Section)
        If temp > 0 Then
            sectioncount = countsections(temp)
        End If
    End If
End Function
Public Property Let keyname(Section As String, key As String, Name As String)
    If entrycount = 0 Then Exit Property
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        If keyexists(temp, key) = True And keyexists(temp, Name) = False Then
            renamekey temp, key, Name
        End If
    End If
End Property
Public Property Let sectionname(Section As String, Name As String)
    If entrycount = 0 Then Exit Property
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        renamesection temp, Name
    End If
End Property
Private Function isroot(text As String) As Boolean
'If Left(text, 2) = "[/" & Right(text, 1) = "]" Then isroot = True Else isroot = False
isroot = text = "[/]"
End Function
Private Function issection(Value As String) As Boolean
On Error Resume Next
    If Left(Value, 1) = "[" And Right(Value, 1) = "]" And stripsection(Value) <> Empty Then issection = True Else issection = False
End Function
Private Function isnotroot(Value As String) As Boolean
isnotroot = Not Value = "[/]" 'If issection(value) = True And isroot(value) = False Then isnotroot = True
End Function
Private Function isvalue(Value As String) As Boolean
On Error Resume Next
    If issection(Value) = False And InStr(Value, "=") > 0 Then isvalue = True Else isvalue = False
End Function
Private Function stripsection(Section As String) As String
On Error Resume Next
    stripsection = Mid(Section, 2, Len(Section) - 2)
End Function
Private Function stripvalue(Value As String) As String
On Error Resume Next
    stripvalue = Right(Value, Len(Value) - InStr(Value, "="))
End Function
Private Function stripname(Value As String) As String
On Error Resume Next
    stripname = Left(Value, InStr(Value, "=") - 1)
End Function
Private Function iscomment(Value As String) As Boolean
On Error Resume Next
    If Left(Value, 1) = "#" Or Left(Value, 1) = "'" Then iscomment = True Else iscomment = False
End Function
Private Sub UserControl_Initialize()
    Dim tempstr() As String
    isloaded = False
    entrycount = 0
End Sub
Public Sub savefile(filename As String)
If entrycount = 0 Then Exit Sub
    Dim tempfile As Long, count As Long
If isloaded = True Then
tempfile = FreeFile
Open filename For Output As tempfile
    For count = 1 To entrycount
        Print #tempfile, String(inifile(count).level, vbTab) & inifile(count).contents
    Next
Close tempfile
Else
    errorcode = err_filenotloaded
End If
End Sub
Public Sub saveoldfile(filename As String)
If entrycount = 0 Then Exit Sub
    Dim temp As Long
temp = FreeFile
If filename Like "?:\*" Then
    Open filename For Output As temp
        saveoldroot temp
    Close temp
End If
End Sub
Private Sub saveoldroot(Filenumber As Long)
If entrycount = 0 Then Exit Sub
    Dim tempstr() As String, count As Long, temp As Long
temp = countrootkeys
If temp > 0 Then
    Print #Filenumber, "[/]"
    enumrootkeys tempstr
    For count = 1 To temp
        Print #Filenumber, tempstr(1, count) & "=" & tempstr(2, count)
    Next
End If
temp = countrootsections
If temp > 0 Then
    enumrootsections tempstr
    For count = 1 To temp
        saveoldsection tempstr(count), Filenumber
    Next
End If
End Sub

Private Sub saveoldsection(Section As String, Filenumber As Long)
If entrycount = 0 Then Exit Sub
    Dim temp As Long, tempstr() As String, count As Long, tempstr2() As String
temp = qualifiedsectionhandle(Section)
If countsections(temp) > 0 Then
    enumsections temp, tempstr
    For count = 1 To countsections(temp)
        saveoldsection Section & "\" & tempstr(count), Filenumber
    Next
End If
Print #Filenumber, "[" & Replace(Section, "\", "/") & "]"
enumkeys temp, tempstr2
For count = 1 To countkeys(temp)
    If countinstancekeys(temp, tempstr2(1, count)) = 1 Then
        Print #Filenumber, tempstr2(1, count) & "=" & tempstr2(2, count)
    Else
        Print #Filenumber, tempstr2(1, count) & "(" & getkeysinstance(temp, counthandlekeys(temp, count)) & ")=" & tempstr2(2, count)
    End If
Next
Print #Filenumber, vbNewLine 'whitespace
End Sub
Public Sub loadoldfile(filename As String)
On Error Resume Next
If FileLen(filename) = 0 Then Exit Sub
    Dim tempfile As Long, currlevel As Long, tempstr As String, continue As Boolean
    entrycount = 0
    currlevel = 0
    tempfile = FreeFile

    isloaded = False
    If Dir(filename) <> Empty Then
        Open filename For Input As tempfile
            Do Until EOF(tempfile)
                Line Input #tempfile, tempstr
                tempstr = Replace(tempstr, vbTab, Empty)
                If tempstr <> Empty And tempstr <> "[/]" Then 'removes blank lines and the [/] from hini2ini converted files
                    entrycount = entrycount + 1
                    If issection(tempstr) And currlevel > 0 Then
                        tempstr = Replace(tempstr, "\", "/")
                        entrycount = entrycount + 1
                        ReDim Preserve inifile(1 To entrycount)
                        inifile(entrycount - 1).level = 0
                        inifile(entrycount - 1).contents = "[/]"
                        currlevel = currlevel - 1
                    Else
                        If entrycount = 1 Then
                            ReDim inifile(1 To 1)
                        Else
                            ReDim Preserve inifile(1 To entrycount)
                        End If
                    End If
                    inifile(entrycount).level = currlevel
                    inifile(entrycount).contents = tempstr
                    If issection(tempstr) Then currlevel = currlevel + 1
                End If
            Loop
            If currlevel > 0 Then
                entrycount = entrycount + 1
                ReDim Preserve inifile(1 To entrycount)
                inifile(entrycount).level = 0
                inifile(entrycount).contents = "[/]"
            End If
        Close tempfile
        isloaded = True
    End If
End Sub

Public Sub loadfile(filename As String, Optional appendtoroot As String = Empty)
On Error Resume Next
If FileLen(filename) = 0 Then Exit Sub
    Dim tempfile As Long, currlevel As Long, tempstr As String, continue As Boolean
    If appendtoroot = Empty Then
        entrycount = 0
        currlevel = 0
    Else
        entrycount = entrycount + 1
        ReDim Preserve inifile(1 To entrycount)
        inifile(entrycount).level = 0
        inifile(entrycount).contents = "[" & appendtoroot & "]"
        currlevel = 1
    End If
    tempfile = FreeFile
    errorcode = 0

    If Dir(filename) <> Empty Then
        Open filename For Input As tempfile
            Do Until EOF(tempfile)
                Line Input #tempfile, tempstr
                tempstr = Replace(tempstr, vbTab, Empty)
                If tempstr <> Empty Then  '[/] sections = evil
                    continue = True ' account for roots going below root level
                    If isroot(tempstr) Then
                        If currlevel - 1 >= 0 Then
                            currlevel = currlevel - 1
                        Else
                            continue = False
                        End If
                    End If
                    If continue = True Then
                        entrycount = entrycount + 1
                        ReDim Preserve inifile(1 To entrycount)
                        inifile(entrycount).level = currlevel
                        If isroot(tempstr) = False And issection(tempstr) = True Then currlevel = currlevel + 1
                        inifile(entrycount).contents = tempstr
                    End If
                End If
            Loop
        Close tempfile
        
        If currlevel > 0 Then 'accounts for missing roots
            For tempfile = currlevel To 1 Step -1
                entrycount = entrycount + 1
                ReDim Preserve inifile(1 To entrycount)
                inifile(entrycount).level = tempfile - 1
                inifile(entrycount).contents = "[/]"
            Next
        End If
        
        isloaded = True
    Else
        errorcode = err_filedoesntexist
    End If
End Sub
Public Sub closefile()
    entrycount = 0
    ReDim inifile(0)
    isloaded = False
End Sub
Private Function findroot(start As Long) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, found As Boolean
found = False
If start > 0 Then
If isloaded And start > 0 Then
    For count = start + 1 To entrycount
        If found = False Then
            If inifile(count).level = inifile(start).level Then
'                If isroot(inifile(count).contents) Then
                    found = True
                    findroot = count
'                End If
'This code is by Techni Myoko, and anyone who claims otherwise is a liar
            End If
        End If
    Next
End If
Else
    findroot = entrycount
End If
End Function
Private Function countkeys(start As Long) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If isvalue(inifile(count).contents) Then
            keys = keys + 1
        End If
    End If
Next
countkeys = keys
End Function
Private Function counthandlekeys(start As Long, Index As Long) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If isvalue(inifile(count).contents) Then
            keys = keys + 1
            If Index = keys Then
                counthandlekeys = count
            End If
        End If
    End If
Next
End Function
Private Function countinstancekeys(start As Long, key As String) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long, temp As Long, temp2 As Long
If start = 0 Then
    temp = entrycount
    temp2 = 0
    start = 1
Else
    temp = findroot(start)
    temp2 = inifile(start).level + 1
End If
For count = start To temp
    If inifile(count).level = temp2 Then
        If isvalue(inifile(count).contents) And LCase(stripname(inifile(count).contents)) = LCase(key) Then
            keys = keys + 1
        End If
    End If
Next
countinstancekeys = keys
End Function
Private Function getkeysinstance(start As Long, handle As Long) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long, temp As String
temp = LCase(stripname(inifile(handle).contents))
For count = start To handle
    If inifile(count).level = inifile(start).level + 1 Then
        If isvalue(inifile(count).contents) And LCase(stripname(inifile(count).contents)) = temp Then
            keys = keys + 1
        End If
    End If
Next
getkeysinstance = keys
End Function
Private Function countsections(start As Long) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
If start > 0 Then
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If isnotroot(inifile(count).contents) Then
            sections = sections + 1
        End If
    End If
Next
Else
    countsections = countrootsections
End If
countsections = sections
End Function
Public Function countrootsections() As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
For count = 1 To entrycount
    If inifile(count).level = 0 Then
    If inifile(count).contents <> Empty Then
        If isnotroot(inifile(count).contents) Then
            sections = sections + 1
        End If
    End If
    End If
Next
countrootsections = sections
End Function
Public Sub enumeratesections(Section As String, STRarray)
If entrycount = 0 Then Exit Sub
    If Section <> Empty Then
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        enumsections temp, STRarray
    End If
Else
    enumrootsections STRarray
End If
End Sub
Public Function enumrootsections(STRarray)
If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
For count = 1 To entrycount
    If inifile(count).level = 0 Then
    If inifile(count).contents <> Empty Then
        If isnotroot(inifile(count).contents) Then
            sections = sections + 1
            ReDim Preserve STRarray(1 To sections)
            STRarray(sections) = stripsection(inifile(count).contents)
        End If
    End If
    End If
Next
End Function
Public Function handlerootsections(Section As String) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long
For count = 1 To entrycount
    If inifile(count).level = 0 Then
        If isnotroot(inifile(count).contents) = True Then
            If LCase(stripsection(inifile(count).contents)) = LCase(Section) Then
                handlerootsections = count
            End If
        End If
    End If
Next
End Function

Private Sub enumsections(start As Long, STRarray)
If entrycount = 0 Then Exit Sub
    Dim count As Long, sections As Long
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If isnotroot(inifile(count).contents) Then
            sections = sections + 1
            ReDim Preserve STRarray(1 To sections)
            STRarray(sections) = stripsection(inifile(count).contents)
        End If
    End If
Next
End Sub
Public Sub enumeratekeys(Section As String, STRarray)
If entrycount = 0 Then Exit Sub
    If Section <> Empty Then
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        enumkeys temp, STRarray
    End If
Else
    enumrootkeys STRarray
End If
End Sub
Public Function enumrootkeys(STRarray)
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
For count = 1 To entrycount
    If inifile(count).level = 0 Then
        If isvalue(inifile(count).contents) Then
            keys = keys + 1
            ReDim Preserve STRarray(1 To 2, 1 To keys) As String
            STRarray(1, keys) = stripname(inifile(count).contents)
            STRarray(2, keys) = stripvalue(inifile(count).contents)
        End If
    End If
Next
End Function
Public Function countrootkeys() As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
For count = 1 To entrycount
    If inifile(count).level = 0 Then
        If isvalue(inifile(count).contents) Then
            keys = keys + 1
        End If
    End If
Next
countrootkeys = keys
End Function
Private Sub enumkeys(start As Long, STRarray)
If entrycount = 0 Then Exit Sub
    Dim count As Long, keys As Long
If start > 0 Then
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If isvalue(inifile(count).contents) Then
            keys = keys + 1
            ReDim Preserve STRarray(1 To 2, 1 To keys) As String
            STRarray(1, keys) = stripname(inifile(count).contents)
            STRarray(2, keys) = stripvalue(inifile(count).contents)
        End If
    End If
Next
Else
enumrootkeys STRarray
End If
End Sub
Private Function sectionexists(start As Long, Section As String) As Boolean
If entrycount = 0 Then Exit Function
        Dim tempstr() As String, count As Long, temp As Long
    If start > 0 Then
        enumsections start, tempstr
        temp = countsections(start)
    Else
        enumrootsections tempstr
        temp = countrootsections
    End If
    sectionexists = False
    For count = 1 To temp
        If LCase(tempstr(count)) = LCase(Section) Then
            sectionexists = True
            Exit For
        End If
    Next
End Function
Public Function sectionatindex(Section As String, Index As Long)
If entrycount = 0 Then Exit Function
        Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        sectionatindex = sectionindex(temp, Index)
    End If
End Function
Private Function sectionindex(start As Long, Index As Long) As String
If entrycount = 0 Or Index = 0 Then Exit Function
        Dim tempstr() As String, temp As Long
    If start > 0 Then
        enumsections start, tempstr
        temp = countsections(start)
    Else
        enumrootsections tempstr
        temp = countrootsections
    End If
    If Index <= temp And temp > 0 Then
        sectionindex = tempstr(Index)
    End If
End Function
Public Function existsancekey(Section As String, key As String, Optional Index As Long = 0) As Boolean
If entrycount = 0 Then Exit Function
        Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    existsancekey = False
    If temp > 0 Then
        existsancekey = keyexists(temp, key, Index)
    End If
End Function
Public Function existancesection(Section As String) As Boolean
If entrycount = 0 Then Exit Function
        Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    existancesection = False
    If temp > 0 Then existancesection = True
End Function
Private Function keyexists(start As Long, key As String, Optional Index As Long = 0) As Boolean
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String, count As Long, count2 As Long
    If start > 0 Then enumkeys start, tempstr Else enumrootkeys tempstr
    keyexists = False
    For count = 1 To countkeys(start)
        If LCase(tempstr(1, count)) = LCase(key) Then
            count2 = count2 + 1
            If count2 = Index Or Index = 0 Then
                keyexists = True
            End If
        End If
    Next
End Function
Public Function keyindex(start As Long, Index As Long, Optional side As Long = 2) As String
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String
    enumkeys start, tempstr
    If Index > 0 And Index <= countkeys(start) Then
        keyindex = tempstr(side, Index)
    End If
End Function
Public Function keysindex(Section As String, key As String, Optional Index As Long = 0) As Long
If entrycount = 0 Then Exit Function
        Dim tempstr() As String, temp As Long, count As Long, count2 As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        enumkeys temp, tempstr
        For count = 1 To countkeys(temp)
            If LCase(tempstr(count)) = LCase(key) Then
                count2 = count2 + 1
                If count2 = Index Or Index = 0 Then
                    keysindex = count
                End If
            End If
        Next
    End If
End Function
Private Function keyhandle(start As Long, key As String, Optional Index As Long = 0) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long, count2 As Long, found As Boolean
found = False
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 And found = False Then
        If isvalue(inifile(count).contents) Then
            If LCase(stripname(inifile(count).contents)) = LCase(key) Then
                count2 = count2 + 1
                If count2 = Index Or Index = 0 Then
                    keyhandle = count
                    found = True
                End If
            End If
        End If
    End If
Next
End Function
Public Function newrootsection(Section As String)
    If sectionexists(0, Section) = False Then
isloaded = True
entrycount = entrycount + 2
If entrycount > 2 Then ReDim Preserve inifile(1 To entrycount)
If entrycount = 2 Then ReDim inifile(1 To entrycount)
inifile(entrycount - 1).level = 0
inifile(entrycount - 1).contents = "[" & Section & "]"
inifile(entrycount).level = 0
inifile(entrycount).contents = "[/]"
End If
End Function
Public Function enummultikey(Section As String, key As String, STRarray) As Long
If entrycount = 0 Then Exit Function
    Dim start As Long, count As Long, temp As Long, count2 As Long, count3 As Long
start = qualifiedsectionhandle(Section)
If start = 0 Then
    temp = 0
    start = 1
Else
    temp = inifile(start).level + 1
End If

count3 = countmultikey(start, key)
If count3 > 0 Then
ReDim Preserve STRarray(1 To count3) As String
For count = start To findroot(start)
    If inifile(count).level = temp Then
        If isvalue(inifile(count).contents) Then
            If LCase(stripname(inifile(count).contents)) = LCase(key) Then
                count2 = count2 + 1
                STRarray(count2) = stripvalue(inifile(count).contents)
            End If
        End If
    End If
Next
End If
enummultikey = count3
End Function
Public Function multikeycount(Optional Section As String, Optional key As String) As Long
If entrycount = 0 Then Exit Function
    multikeycount = countmultikey(qualifiedsectionhandle(Section), key)
End Function
Private Function countmultikey(start As Long, key As String) As Long
If entrycount = 0 Then Exit Function
    On Error Resume Next
Dim count As Long, temp As Long, count2 As Long
If start = 0 Then
    temp = 0
    start = 1
Else
    temp = inifile(start).level + 1
End If
For count = start To findroot(start)
    If inifile(count).level = temp Then
        If isvalue(inifile(count).contents) Then
            If LCase(stripname(inifile(count).contents)) = LCase(key) Then
                count2 = count2 + 1
            End If
        End If
    End If
Next
countmultikey = count2
End Function
Private Function getkey(start As Long, key As String, Optional Index As Long = 0) As String
If entrycount = 0 Then Exit Function
    Dim count As Long, count2 As Long, temp As Long, found As Boolean
    'This code is by Techni Myoko, and anyone who claims otherwise is a liar
found = False
If start = 0 Then
    temp = 0
    start = 1
Else
    temp = inifile(start).level + 1
End If
For count = start To findroot(start)
    If inifile(count).level = temp And found = False Then
        If isvalue(inifile(count).contents) Then
            If LCase(stripname(inifile(count).contents)) = LCase(key) Then
                count2 = count2 + 1
                If Index = 0 Or count2 = Index Then
                    getkey = stripvalue(inifile(count).contents)
                    found = True
                End If
            End If
        End If
    End If
Next
End Function
Public Sub createkey(Section As String, key As String, Value As String, Optional Index As Long = 0)
Dim temp As Long
temp = qualifiedsectionhandle(Section)
If temp > 0 Then
    temp = findroot(temp)
    If keyexists(temp, key, Index) = False Then
        insert temp, inifile(temp).level + 1, key & "=" & Value
    Else
        setkeycontents Section, key, Value, Index
    End If
End If
End Sub
Public Function Line(Index As Long) As String
Line = inifile(Index).level & ") " & inifile(Index).contents
End Function
Public Function LineCount() As Long
LineCount = entrycount
End Function
Private Sub setkey(start As Long, key As String, Value As String, Optional Index As Long = 0)
Dim count As Long, count2 As Long, found As Boolean
found = False
If keyexists(start, key) = True Then
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 And found = False Then
        If isvalue(inifile(count).contents) Then
            If LCase(stripname(inifile(count).contents)) = LCase(key) Then
                count2 = count2 + 1
                If Index = 0 Or count2 = Index Then
                    inifile(count).contents = stripname(inifile(count).contents) & "=" & Value
                    found = True
                End If
            End If
        End If
    End If
Next
End If
End Sub
Private Function sectionhandle(start As Long, Section As String) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long
If start > 0 Then
For count = start To findroot(start)
    If inifile(count).level = inifile(start).level + 1 Then
        If issection(inifile(count).contents) Then
            If LCase(stripsection(inifile(count).contents)) = LCase(Section) Then
                sectionhandle = count
            End If
        End If
    End If
Next
Else
    sectionhandle = handlerootsections(Section)
End If
End Function
Public Function qualifiedsectionhandle(Section As String) As Long
If entrycount = 0 Then Exit Function
    Dim tempstr() As String, count As Long, temp As Long, exists As Boolean
If Section <> Empty Then
tempstr = Split(Section, "\")
exists = sectionexists(0, tempstr(0))
If exists Then
    temp = handlerootsections(tempstr(0))
    For count = 1 To UBound(tempstr)
        exists = exists And sectionexists(temp, tempstr(count))
        If exists = True Then
            temp = sectionhandle(temp, tempstr(count))
        End If
    Next
End If
If exists = True Then qualifiedsectionhandle = temp
End If
End Function
Private Sub renamekey(start As Long, key As String, Name As String, Optional Index As Long = 0, Optional newindex As Long = 0)
If entrycount = 0 Then Exit Sub
    Dim temp As Long
temp = keyhandle(start, key, Index)
If temp > 0 And keyexists(start, key, newindex) = False Then
    inifile(temp).contents = Name & "=" & stripvalue(inifile(temp).contents)
End If
End Sub
Public Sub deletekey(Section As String, key As String, Optional Index As Long = 0)
 If entrycount = 0 Then Exit Sub
       Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        temp = keyhandle(temp, key, Index)
        removerange temp, temp
    End If
End Sub
Public Sub deletesection(Section As String)
If entrycount = 0 Then Exit Sub
        Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        removesection temp
    End If
End Sub
Private Sub removesection(start As Long)
If entrycount = 0 Then Exit Sub
        removerange start, findroot(start)
End Sub
Private Sub renamesection(start As Long, Section As String)
If entrycount = 0 Then Exit Sub
        If sectionexists(start, Section) = False And Section <> "/" Then
        inifile(start).contents = "[" & Replace(Section, "\", "/") & "]"
    End If
End Sub
Public Function concatenate(STRarray, delimeter As String) As String
Dim temp As String, count As Long
For count = LBound(STRarray) To UBound(STRarray)
    If temp <> Empty Then temp = temp & delimeter
    temp = temp & STRarray(count)
Next
End Function
Public Sub creationsection(Section As String)
Dim tempstr() As String, tempstr2() As Long, count As Long
tempstr = Split(Section, "\")
ReDim Preserve tempstr2(UBound(tempstr)) As Long
For count = 0 To UBound(tempstr)
    If count = 0 Then
        If sectionexists(0, tempstr(0)) = False Then
            newrootsection tempstr(0)
        End If
        tempstr2(0) = handlerootsections(tempstr(0))
    Else
        If sectionexists(tempstr2(count - 1), tempstr(count)) = False Then
            createsection tempstr2(count - 1), tempstr(count)
        End If
        tempstr2(count) = sectionhandle(tempstr2(count - 1), tempstr(count))
    End If
Next
End Sub
Private Sub createsection(ByVal start As Long, Section As String)
    Dim finish As Long
    Section = Replace(Section, "\", "/")
    If start > 0 Then
    finish = findroot(start)
    If sectionexists(start, Section) = False And Section <> "/" And Section <> Empty Then
        insert finish, inifile(start).level + 1, "[/]"
        insert finish, inifile(start).level + 1, "[" & Section & "]"
    End If
    Else
        newrootsection Section
    End If
End Sub
Private Sub insert(start As Long, level As Long, contents As String)
Dim count As Long
entrycount = entrycount + 1
ReDim Preserve inifile(1 To entrycount)
For count = entrycount - 1 To start Step -1
    inifile(count + 1).contents = inifile(count).contents
    inifile(count + 1).level = inifile(count).level
Next
inifile(start).contents = contents
inifile(start).level = level
End Sub
Private Sub removerange(top As Long, Bottom As Long)
If entrycount = 0 Then Exit Sub
        Dim range As Long, count As Long
    If top > 0 And Bottom > 0 And top <= entrycount And Bottom <= entrycount Then
    range = Bottom + 1 - top
    entrycount = entrycount - range
    For count = top To entrycount
        inifile(count).contents = inifile(count + range).contents
        inifile(count).level = inifile(count + range).level
    Next
    If entrycount > 0 Then ReDim Preserve inifile(1 To entrycount) Else ReDim inifile(entrycount)
    End If
End Sub
Public Function findnext(text As String, ByVal start As Long, char As String) As Long
    start = start + 1
    Do Until LCase(Mid(text, start, Len(char))) = LCase(char) Or start = Len(text)
        start = start + 1
    Loop
    findnext = start
End Function
Public Function findprev(text As String, ByVal start As Long, char As String) As Long
    start = start - 1
    Do Until LCase(Mid(text, start, Len(char))) = LCase(char) Or start <= 1
        start = start - 1
    Loop
    findprev = start
End Function
Private Sub ReDimPreserve(ByRef count As Long, STRarray() As entry, level As Long, contents As String)
    count = count + 1
    ReDim Preserve STRarray(1 To count) As entry
    STRarray(count).level = level
    STRarray(count).contents = contents
End Sub
Public Sub loadxmlfile(filename As String) 'I hate xml
    Dim xmlfile() As entry, xmlentries As Long, xml As String
    Dim count As Long, count2 As Long, count3 As Long, count4 As Long, temp As String
    Dim X1 As String, X2 As String, currlevel As Long
    xml = loadwholefile(filename)
    If xml <> Empty Then 'loads an xml into a hierarchical structure
        count = findnext(xml, 0, "<")
        Do Until InStr(Right(xml, Len(xml) - count), "<") = 0 Or count >= Len(xml)
            count2 = findnext(xml, count, ">") + 1
            temp = Trim(Replace(Replace(Mid(xml, count, count2 - count), vbNewLine, Empty), vbTab, Empty))
            If InStr(temp, "</") > 0 Then
                X1 = Left(temp, InStr(temp, "<") - 1) 'contents of tag
                X2 = Right(temp, Len(temp) - InStr(temp, "<") + 1) 'end tag
                If X1 <> Empty Then ReDimPreserve xmlentries, xmlfile, currlevel, X1
                currlevel = currlevel - 1
                ReDimPreserve xmlentries, xmlfile, currlevel, X2
            Else
                If Right(temp, 2) = "/>" Or Right(temp, 2) = "?>" Then
                    ReDimPreserve xmlentries, xmlfile, currlevel, temp
                Else
                    ReDimPreserve xmlentries, xmlfile, currlevel, temp
                    currlevel = currlevel + 1
                End If
            End If
            count = count2
            If count > Len(xml) Then count = Len(xml)
        Loop
    End If
    
    entrycount = 0
    currlevel = 0
    For count = 1 To xmlentries 'convert hierarchical structure to hini
        'entrycount = entrycount + 1
            Do Until InStr(xmlfile(count).contents, "  ") = 0
                xmlfile(count).contents = Replace(xmlfile(count).contents, "  ", " ")
            Loop
            Do Until InStr(xmlfile(count).contents, " =") = 0
                xmlfile(count).contents = Replace(xmlfile(count).contents, " =", "=")
            Loop
            Do Until InStr(xmlfile(count).contents, "= ") = 0
                xmlfile(count).contents = Replace(xmlfile(count).contents, "= ", "=")
            Loop
        If Left(xmlfile(count).contents, 1) = "<" And Left(xmlfile(count).contents, 2) <> "</" Then
            'isnt and end tag, can have parameters
            count3 = InStr(xmlfile(count).contents, " ")
            If count3 = 0 Then 'no parameters
                ReDimPreserve entrycount, inifile, currlevel, "[" & Replace(Mid(xmlfile(count).contents, 2, Len(xmlfile(count).contents) - 2), "/", Empty) & "]"
                currlevel = currlevel + 1
                If Right(Mid(xmlfile(count).contents, 2, Len(xmlfile(count).contents) - 2), 2) = "/>" Then
                    currlevel = currlevel - 1
                    ReDimPreserve entrycount, inifile, currlevel, "[/]"
                End If
            Else
                ReDimPreserve entrycount, inifile, currlevel, "[" & Mid(xmlfile(count).contents, 2, count3 - 2) & "]"
                count2 = 0
                currlevel = currlevel + 1
                Do Until InStr(Right(xmlfile(count).contents, Len(xmlfile(count).contents) - count2), "=") = 0
                    count2 = findnext(xmlfile(count).contents, count3, "=")
                    count3 = findprev(xmlfile(count).contents, count2, " ")
                    count4 = findnext(xmlfile(count).contents, count2 + 2, """")
                    temp = Mid(xmlfile(count).contents, count3, count4 - count3 + 1)
                    count3 = count2 + 1
                    ReDimPreserve entrycount, inifile, currlevel, Replace(Left(temp, Len(temp) - 1), "=""", "=")
                Loop
            End If
            If Right(xmlfile(count).contents, 2) = "/>" Or Right(xmlfile(count).contents, 2) = "?>" Then
                currlevel = currlevel - 1
                ReDimPreserve entrycount, inifile, currlevel, "[/]"
            End If
        Else
            If Left(xmlfile(count).contents, 2) = "</" Then
                currlevel = currlevel - 1
                ReDimPreserve entrycount, inifile, currlevel, "[/]"
            Else
                ReDimPreserve entrycount, inifile, currlevel, "Node=" & xmlfile(count).contents
            End If
        End If
    Next
    isloaded = True
    removedoubles
End Sub
Public Sub removedoubles()
If entrycount = 0 Then Exit Sub
         Dim count As Long, count2 As Long
     For count = 1 To entrycount
        If isnotroot(inifile(count).contents) Then
            removedoublessection count
            'This code is by Techni Myoko, and anyone who claims otherwise is a liar
            'removedoublekey count ' not needed with multikey
        End If
     Next
End Sub
Private Sub removedoublekey(start As Long)
If entrycount = 0 Then Exit Sub
        Dim count As Long, count2 As Long, count3 As Long
    Dim root As Long
    root = findroot(start) - 1
    For count = start + 1 To root
        If inifile(count).level = inifile(start).level + 1 Then
            If isvalue(inifile(count).contents) Then
                count3 = 0
                For count2 = count + 1 To root
                    If inifile(count2).level = inifile(count).level Then
                        If isvalue(inifile(count2).contents) Then
                            If LCase(stripname(inifile(count2).contents)) = LCase(stripname(inifile(count).contents)) Then
                                count3 = count3 + 1
                                inifile(count2).contents = Left(inifile(count2).contents, InStr(inifile(count2).contents, "=") - 1) & "(" & count3 & ")=" & Right(inifile(count2).contents, Len(inifile(count2).contents) = InStr(inifile(count2).contents, "="))
                            End If
                        End If
                    End If
                Next
                If count3 > 0 Then
                    inifile(count).contents = Left(inifile(count).contents, InStr(inifile(count).contents, "=") - 1) & "(0)=" & Right(inifile(count).contents, Len(inifile(count).contents) = InStr(inifile(count).contents, "="))
                End If
            End If
        End If
    Next
End Sub
Private Sub removedoublessection(start As Long)
If entrycount = 0 Then Exit Sub
        Dim count As Long, count2 As Long, count3 As Long
    Dim root As Long
    root = findroot(start) - 1
   ' MsgBox root
    For count = start + 1 To root
        If inifile(count).level = inifile(start).level + 1 Then
            If isnotroot(inifile(count).contents) Then
                count3 = 0
                For count2 = count + 1 To root
                    If inifile(count2).level = inifile(count).level Then
                        If isnotroot(inifile(count2).contents) Then
                            If LCase(stripsection(inifile(count2).contents)) = LCase(stripsection(inifile(count).contents)) Then
                                count3 = count3 + 1
                                inifile(count2).contents = Left(inifile(count2).contents, Len(inifile(count2).contents) - 1) & "(" & count3 & ")]"
                            End If
                        End If
                    End If
                Next
                If count3 > 0 Then
                    inifile(count).contents = Left(inifile(count).contents, Len(inifile(count).contents) - 1) & "(0)]"
                End If
            End If
        End If
    Next
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 480
UserControl.Height = UserControl.Width
End Sub
Public Property Let Tag(text As String)
    stag = text
End Property
Public Property Get Tag() As String
    Tag = stag
End Property
'This code is by Techni Myoko, and anyone who claims otherwise is a liar

