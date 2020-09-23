Attribute VB_Name = "ListViewHandling"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const LVM_FIRST = &H1000
Public Sub autosizeall(lview As ListView)
If lview.ListItems.count > 0 Then
Dim count As Integer
For count = 1 To lview.ColumnHeaders.count
    AutoSizeColumnHeader lview, lview.ColumnHeaders.item(count)
Next
autoalign lview
End If
End Sub
Public Sub AutoSizeColumnHeader(lview As ListView, column As ColumnHeader, Optional ByVal SizeToHeader As Boolean = True)
On Error Resume Next
    Dim l As Long
    If SizeToHeader Then l = -2 Else l = -1
    Call SendMessage(lview.hwnd, LVM_FIRST + 30, column.Index - 1, l)
End Sub
Public Function selecteditem(lst As ListView)
    On Error Resume Next
    selecteditem = 0
    selecteditem = lst.selecteditem.Index
End Function
Public Sub resizecolumnheaders(lview As ListView)
On Error Resume Next
Dim temp As Integer
If lview.ListItems.count > 0 Then
    For temp = 1 To lview.ColumnHeaders.count
        AutoSizeColumnHeader lview, lview.ColumnHeaders.item(temp)
    Next
End If
End Sub

Public Sub autoalign(lview As ListView)
Dim count As Long, count2 As Long, foundnonnumeric As Boolean
For count = 2 To lview.ColumnHeaders.count
    foundnonnumeric = False
    For count2 = 1 To lview.ListItems.count
        If isnumeric2(getitem(lview, count, count2)) = False Then foundnonnumeric = True
    Next
    If foundnonnumeric = True Then lview.ColumnHeaders.item(count).Alignment = lvwColumnLeft
    If foundnonnumeric = False Then lview.ColumnHeaders.item(count).Alignment = lvwColumnRight
Next
lview.Refresh
End Sub
Public Function getitem(lview As ListView, x As Long, y As Long)
    If x = 1 Then
        getitem = lview.ListItems.item(y).text
    Else
        getitem = lview.ListItems.item(y).SubItems(x - 1)
    End If
End Function
Public Function isnumeric2(text As String) As Boolean
    isnumeric2 = IsNumeric(Replace(Replace(text, ".", ""), "-", ""))
End Function
Public Sub additem(lst As ListView, align As Boolean, ParamArray Items() As Variant)
    Dim temp As Long
    lst.ListItems.Add , , Items(0)
    For temp = 1 To UBound(Items)
        lst.ListItems(lst.ListItems.count).SubItems(temp) = Items(temp)
    Next
    If align Then autosizeall lst
End Sub
Public Sub selectall(lst As ListView)
    Dim temp As Long
    For temp = 1 To lst.ListItems.count
        lst.ListItems.item(temp).Checked = True
    Next
    lst.Refresh
End Sub
Public Sub removedoubles(lstmain As ListView, item As Long)
'Remove doubles of the resolved urls
Dim spoth As Long, spoth2 As Long
For spoth = 1 To lstmain.ListItems.count 'To 1
    DoEvents
    If spoth <= lstmain.ListItems.count Then
    For spoth2 = lstmain.ListItems.count To spoth + 1 Step -1
            If StrComp(getitem(lstmain, item, spoth), getitem(lstmain, item, spoth2), vbTextCompare) = 0 Then lstmain.ListItems.Remove spoth2
    Next
    End If
Next
End Sub
