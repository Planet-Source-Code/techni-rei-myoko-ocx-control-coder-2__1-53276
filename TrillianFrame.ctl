VERSION 5.00
Begin VB.UserControl TrillianFrame 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   ControlContainer=   -1  'True
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ToolboxBitmap   =   "TrillianFrame.ctx":0000
   Begin VB.Label lblmain 
      BackColor       =   &H00E7A27B&
      Caption         =   "Trillian Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   1335
   End
   Begin VB.Image imgmain 
      Height          =   15
      Index           =   0
      Left            =   0
      Picture         =   "TrillianFrame.ctx":0312
      Top             =   240
      Width           =   1440
   End
   Begin VB.Shape Shpmain 
      BorderColor     =   &H00E7A27B&
      BorderWidth     =   4
      Height          =   3255
      Index           =   1
      Left            =   45
      Top             =   285
      Width           =   2175
   End
   Begin VB.Shape Shpmain 
      BackColor       =   &H00E7A27B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E7A27B&
      FillColor       =   &H00E7A27B&
      Height          =   225
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   2220
   End
   Begin VB.Shape Shpmain 
      BackColor       =   &H00E7A27B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E7A27B&
      FillColor       =   &H00E7A27B&
      Height          =   15
      Index           =   2
      Left            =   15
      Top             =   240
      Width           =   2220
   End
End
Attribute VB_Name = "TrillianFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const dark_blue As Long = 14053681
Const light_blu As Long = 15180411
Const dark_grey As Long = 14074813

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblmain.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
End Sub

Private Sub UserControl_Resize()
If UserControl.Height > 315 Then
    Shpmain(0).Width = UserControl.Width / 15 - 2
    Shpmain(2).Width = Shpmain(0).Width
    lblmain.Width = Shpmain(0).Width - 2
    Shpmain(1).Width = Shpmain(0).Width - 3
    Shpmain(1).Height = UserControl.Height / 15 - 21
End If
End Sub

Public Property Let Caption(text As String)
    lblmain.Caption = text
End Property
Public Property Get Caption() As String
    Caption = lblmain.Caption
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", lblmain.Caption, UserControl.Name
End Sub

