VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Control Coder"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6360
   Icon            =   "clerks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstbox 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "clerks.frx":0E42
      Left            =   360
      List            =   "clerks.frx":0E44
      TabIndex        =   12
      Top             =   4380
      Visible         =   0   'False
      Width           =   2655
   End
   Begin OCXCoder.TrillianFrame trlmain 
      Height          =   2055
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      Caption         =   "Controls"
      Begin VB.CommandButton cmdmain 
         Caption         =   "Set"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   13
         Top             =   45
         Width           =   615
      End
      Begin VB.CommandButton cmdmain 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   2
         Top             =   45
         Width           =   615
      End
      Begin OCXCoder.TrillianFrame trlmain 
         Height          =   750
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1323
         Caption         =   "Property name"
         Begin VB.TextBox txtmain 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2655
         End
      End
      Begin OCXCoder.TrillianFrame trlmain 
         Height          =   750
         Index           =   1
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1323
         Caption         =   "Type of variable"
         Begin VB.ComboBox cbomain 
            Height          =   315
            Index           =   1
            ItemData        =   "clerks.frx":0E46
            Left            =   120
            List            =   "clerks.frx":0E48
            TabIndex        =   9
            Top             =   360
            Width           =   2655
         End
      End
      Begin OCXCoder.TrillianFrame trlmain 
         Height          =   750
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1323
         Caption         =   "Container of variable"
         Begin VB.PictureBox picdummydropdown 
            BackColor       =   &H00FFFFFF&
            HasDC           =   0   'False
            Height          =   315
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   2595
            TabIndex        =   14
            Top             =   360
            Width           =   2655
            Begin VB.CommandButton cmdmain 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Webdings"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2340
               TabIndex        =   15
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtmain 
               BorderStyle     =   0  'None
               Height          =   285
               HideSelection   =   0   'False
               Index           =   2
               Left            =   15
               TabIndex        =   16
               Top             =   15
               Width           =   2400
            End
         End
      End
      Begin OCXCoder.TrillianFrame trlmain 
         Height          =   750
         Index           =   3
         Left            =   3120
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1323
         Caption         =   "Default Value"
         Begin VB.TextBox txtmain 
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton cmdmain 
      Height          =   225
      Index           =   3
      Left            =   6000
      TabIndex        =   4
      ToolTipText     =   "Copy"
      Top             =   8040
      Width           =   225
   End
   Begin VB.TextBox txtmain 
      Height          =   3615
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4680
      Width           =   6135
   End
   Begin MSComctlLib.ListView lstmain 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Container"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Default"
         Object.Width           =   2646
      EndProperty
   End
   Begin OCXCoder.Hini Hinimain 
      Left            =   5760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileop 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnufileop 
         Caption         =   "&Load"
         Index           =   1
      End
      Begin VB.Menu mnufileop 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnufileop 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnueditop 
         Caption         =   "&Remove"
         Index           =   0
      End
      Begin VB.Menu mnueditop 
         Caption         =   "Remove &All"
         Index           =   1
      End
      Begin VB.Menu mnueditop 
         Caption         =   "&Generate All"
         Index           =   2
      End
      Begin VB.Menu mnueditop 
         Caption         =   "&Copy"
         Index           =   3
      End
      Begin VB.Menu mnueditop 
         Caption         =   "&Cut"
         Index           =   4
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filename As String, properties As String, writeprops As String, readprops As String, initprops As String

Private Sub cbomain_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is < 48 'GNDN
        Case Else: AutoComplete cbomain(Index), cbomain(Index)
    End Select
End Sub

Private Sub Form_Load()
AddBuiltInTypes Hinimain
SeedList cbomain(1), "Types", Hinimain, True
End Sub

Private Sub lstbox_Click()
    Dim temp As Long
    temp = txtmain(2).SelStart
    txtmain(2) = ForceAutoComplete(txtmain(2), lstbox.List(lstbox.ListIndex))
    txtmain(2).SelStart = getNextDel(txtmain(2), temp)
End Sub
Private Sub cmdmain_Click(Index As Integer)
On Error Resume Next
If Index = 0 Or Index = 4 Then
    If SubExists(txtmain(0), Hinimain) Then
        MsgBox txtmain(0) & " exists already. A second copy would cause a conflict", vbCritical, "Duplicate sub routine"
        Exit Sub
    End If
End If
Select Case Index
    Case 0 'Add
        lstmain.ListItems.Add , , txtmain(0)
        With lstmain.ListItems.item(lstmain.ListItems.count)
            .SubItems(1) = cbomain(1)
            .SubItems(2) = txtmain(2)
            .SubItems(3) = txtmain(3)
        End With
        RefreshList
    Case 1 'hide/show dropdown
        lstbox.Refresh
        lstbox.Visible = Not lstbox.Visible
        If lstbox.Visible And lstbox.ListCount = 0 Then
            SeedList lstbox, "Variables", Hinimain, False
        End If
        If lstbox.Visible Then lstbox.SetFocus
    Case 2 'Generate
        properties = GenProperties(lstmain) ''Generate Get, Let properties
        readprops = GenReadProperties(lstmain) 'Generate UserControl_ReadProperties
        writeprops = GenWriteProperties(lstmain) 'Generate UserControl_WriteProperties
        initprops = GenInitProperties(lstmain) 'Generate initialize event
        txtmain(4) = properties & readprops & vbNewLine & vbNewLine & writeprops & vbNewLine & vbNewLine & initprops
    Case 3 'Copy
        Clipboard.clear
        Clipboard.SetText txtmain(4)
    Case 4 'Set
        With lstmain.selecteditem
            .text = txtmain(0)
            .SubItems(1) = cbomain(1)
            .SubItems(2) = txtmain(2)
            .SubItems(3) = txtmain(3)
        End With
        RefreshList
End Select
End Sub
Public Function GenInitProperties(lstmain As ListView) As String
    Dim temp As Long, tempstr As String
    tempstr = "Private Sub UserControl_InitProperties()" & vbNewLine
    For temp = 1 To lstmain.ListItems.count
        With lstmain.ListItems.item(temp)
            tempstr = tempstr & vbTab & .text & " = " & .SubItems(3) & vbNewLine
        End With
    Next
    GenInitProperties = tempstr & "End Sub"
End Function
Public Sub RefreshList()
removedoubles lstmain, 0
autosizeall lstmain
End Sub
Public Function GenReadProperties(lstmain As ListView) As String
        Dim temp As Long, tempstr As String
        tempstr = "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine
        For temp = 1 To lstmain.ListItems.count
            With lstmain.ListItems.item(temp)
                tempstr = tempstr & vbTab & .text & " = PropBag.ReadProperty(" & """" & .text & """" & ", " & .SubItems(3) & ")" & vbNewLine
            End With
        Next
        GenReadProperties = tempstr & "End Sub"
End Function
Public Function GenWriteProperties(lstmain As ListView) As String
        Dim temp As Long, tempstr As String
        tempstr = "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)" & vbNewLine
        For temp = 1 To lstmain.ListItems.count
            With lstmain.ListItems.item(temp)
                tempstr = tempstr & vbTab & "PropBag.WriteProperty " & """" & .text & """" & ", " & .SubItems(2) & ", " & .SubItems(3) & vbNewLine
            End With
        Next
        GenWriteProperties = tempstr & "End Sub"
End Function
Public Function GenProperties(lstmain As ListView) As String
    Dim temp As Long, tempstr As String
        For temp = 1 To lstmain.ListItems.count
            With lstmain.ListItems.item(temp)
                tempstr = tempstr & generate(True, .text, .SubItems(1), .SubItems(2)) & vbNewLine & vbNewLine & generate(False, .text, .SubItems(1), .SubItems(2)) & vbNewLine & vbNewLine
            End With
        Next
        GenProperties = tempstr
End Function
Public Function generate(letorget As Boolean, Name As String, typ As String, variable As String) As String
    generate = "Public Property " & IIf(letorget, "Let ", "Get ") & Name & IIf(letorget, "(temp as " & typ & ")", "() as " & typ) & vbNewLine & vbTab & IIf(letorget, variable, Name) & " = " & IIf(letorget, "temp", variable) & vbNewLine & "End Property"
End Function

Private Sub lstbox_KeyPress(KeyAscii As Integer)
lstbox.Visible = False
End Sub

Private Sub lstbox_LostFocus()
    lstbox.Visible = False
End Sub

Private Sub lstbox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lstbox.Visible = False
End Sub

Private Sub lstmain_ItemClick(ByVal item As MSComctlLib.ListItem)
    With item
        txtmain(0).text = .text
        cbomain(1).text = .SubItems(1)
        txtmain(2).text = .SubItems(2)
        txtmain(3).text = .SubItems(3)
    End With
End Sub

Private Sub lstmain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstmain.ListItems.count > 0 And Button = vbRightButton Then mnuedit_Click: PopupMenu mnuedit
End Sub

Private Sub mnuabout_Click()
MsgBox "This version came to be cause a certain (joviak) user seemed to beleive taking my code, changing the locations of the command buttons, and reposting it as his own was ok (even though PSC admins have been deleting his submissions for copying). So I decided to make some additions so complex, that if he stole no one would possibly beleive he did it.", vbInformation, "ActiveX Control Coder"
MsgBox "VBAssimilator, which assimilates your VB code, adding controls, variables, types, subs" & vbNewLine & "IntelliSense Clone which uses the VBAssimilator to autocomplete text" & vbNewLine & "VBFileED, which edits subs within VB code" & vbNewLine & "Extensive use of my Hierarchical INI database", vbInformation, "New additions include"
End Sub

Private Sub mnuedit_Click()
Dim temp As Long, temp2 As Boolean
temp2 = lstmain.ListItems.count > 0
For temp = 0 To mnueditop.UBound
    mnueditop(temp).Enabled = temp2
Next
End Sub

Private Sub mnueditop_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0: lstmain.ListItems.Remove lstmain.selecteditem.Index 'Remove selected
    Case 1: lstmain.ListItems.clear 'remove all
    Case 2
        cmdmain_Click 2 'generate all
        cmdmain_Click 3 'Copy All
    Case 3 'Copy selected
        With lstmain.selecteditem
            Clipboard.clear
            Clipboard.SetText generate(True, .text, .SubItems(1), .SubItems(2)) & vbNewLine & vbNewLine & generate(False, .text, .SubItems(1), .SubItems(2))
        End With
    Case 4 'Cut selected
        mnueditop_Click 3
        mnueditop_Click 0
End Select
End Sub
Public Sub EditFile(filename As String, Optional newfile As String)
    Dim temp As String, tempfile As String
    temp = Hinimain.loadwholefile(filename)
    cmdmain_Click 2
    temp = SetFunction(temp, "UserControl_WriteProperties", writeprops)
    temp = SetFunction(temp, "UserControl_ReadProperties", readprops)
    temp = SetFunction(temp, "UserControl_InitProperties", initprops)
    temp = temp & vbNewLine & vbNewLine & properties
    tempfile = FreeFile
    Open IIf(Len(newfile) > 0, newfile, filename) For Output As tempfile
        Print #tempfile, temp
    Close tempfile
End Sub
Private Sub mnufileop_Click(Index As Integer)
    Dim temp As String, tempfile As Long
    Select Case Index
        Case 0 'Save
            If SubExists("UserControl_WriteProperties", Hinimain) Or SubExists("UserControl_ReadProperties", Hinimain) Or SubExists("UserControl_InitProperties", Hinimain) Then
                EditFile filename
            Else
                tempfile = FreeFile
                Open filename For Append As tempfile
                    Print #tempfile, txtmain(4)
                Close tempfile
            End If
        Case 1 'Load
            InitOpen "Usercontrols (*.ctl)" & Chr(0) & "*.ctl", "Load a usercontrol"
            temp = Open_File(Me.hwnd)
            If Len(temp) > 0 Then
                mnufileop_Click 2
                cbomain(1).clear
                filename = temp
                ScanFile filename, Hinimain
                SeedList cbomain(1), "Types", Hinimain, True
                mnufileop(0).Enabled = True
                mnufileop(2).Enabled = True
            End If
        Case 2 'Close
            lstbox.clear
            lstmain.ListItems.clear
            cbomain(1).clear
            Hinimain.closefile
            AddBuiltInTypes Hinimain
            SeedList cbomain(1), "Types", Hinimain, True
            mnufileop(0).Enabled = False
            mnufileop(2).Enabled = False
        Case 3 'Exit
            Unload Me
            End
    End Select
    lstbox.Refresh
End Sub

Private Sub txtmain_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 2 Then
        If KeyCode = 40 Or KeyCode = 38 Then
            lstbox.Visible = True
            lstbox.SetFocus
            SendKeys IIf(KeyCode = 38, "{up}", "{down}")
        Else
            IntelliSense txtmain(2), KeyCode, Hinimain, lstbox
        End If
    End If
End Sub
