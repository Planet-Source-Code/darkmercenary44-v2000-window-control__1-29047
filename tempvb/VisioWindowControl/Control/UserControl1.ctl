VERSION 5.00
Begin VB.UserControl v2000Window 
   BackColor       =   &H8000000B&
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   ControlContainer=   -1  'True
   MaskColor       =   &H8000000A&
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   2490
   ScaleWidth      =   4380
   ToolboxBitmap   =   "UserControl1.ctx":0014
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E4C0C0&
      FillColor       =   &H00E4C0C0&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   90
      ScaleHeight     =   270
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   1245
      Width           =   4110
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3810
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   30
         Width           =   240
         Begin VB.Image Image1 
            Height          =   135
            Left            =   45
            Picture         =   "UserControl1.ctx":0326
            Top             =   60
            Width           =   150
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   2
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   210
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   3
            Visible         =   0   'False
            X1              =   225
            X2              =   225
            Y1              =   210
            Y2              =   0
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3555
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   30
         Width           =   240
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   4
            Visible         =   0   'False
            X1              =   225
            X2              =   225
            Y1              =   210
            Y2              =   0
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   5
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   210
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   6
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   7
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Image Image2 
            Height          =   135
            Left            =   45
            Picture         =   "UserControl1.ctx":0800
            Top             =   60
            Width           =   150
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3300
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   30
         Width           =   240
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   8
            Visible         =   0   'False
            X1              =   225
            X2              =   225
            Y1              =   210
            Y2              =   0
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   9
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   210
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   10
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line ButBorder 
            BorderColor     =   &H00800000&
            Index           =   11
            Visible         =   0   'False
            X1              =   0
            X2              =   225
            Y1              =   210
            Y2              =   210
         End
         Begin VB.Image Image3 
            Height          =   135
            Left            =   45
            Picture         =   "UserControl1.ctx":0CDA
            Top             =   60
            Width           =   150
         End
      End
      Begin VB.Line TitleDetailHi 
         BorderColor     =   &H80000009&
         Index           =   2
         X1              =   1785
         X2              =   3210
         Y1              =   210
         Y2              =   210
      End
      Begin VB.Line TitleDetail 
         Index           =   2
         X1              =   1785
         X2              =   3210
         Y1              =   195
         Y2              =   195
      End
      Begin VB.Line TitleDetailHi 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   1785
         X2              =   3210
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line TitleDetail 
         Index           =   1
         X1              =   1785
         X2              =   3210
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line TitleDetailHi 
         BorderColor     =   &H80000009&
         Index           =   0
         X1              =   1785
         X2              =   3210
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Label TitleBarCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   540
      End
      Begin VB.Line TitleDetail 
         Index           =   0
         X1              =   1785
         X2              =   3210
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.Image Image6 
      Height          =   165
      Left            =   1005
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Image Image4 
      Height          =   585
      Left            =   3780
      Top             =   345
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   4200
      Top             =   2340
      Width           =   135
   End
   Begin VB.Line Border 
      Index           =   1
      X1              =   525
      X2              =   2985
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Line Border 
      Index           =   2
      X1              =   4365
      X2              =   4365
      Y1              =   105
      Y2              =   2235
   End
   Begin VB.Line Border 
      Index           =   0
      X1              =   0
      X2              =   3840
      Y1              =   2475
      Y2              =   2475
   End
End
Attribute VB_Name = "v2000Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Public Enum Types

Public Enum TBPos
    Top = 0
    Bottom = 1
    Left = 2
    Right = 3
End Enum

Private HasGotFocus As Boolean
Private WithEvents pSniff As SmartSubClass
Attribute pSniff.VB_VarHelpID = -1

Private Const WM_ACTIVATE = &H6

'Default Property Values:
Const m_def_ButtonDisableColor = &H8000000B
Const m_def_TitleBarLayout = 0
Const m_def_Resizable = True
Const m_def_ControlsBorderColor = &H800000
Const m_def_ControlsBackColor = &HE4C0C0
Const m_def_ControlsHoverColor = &HDEB1B1
Const m_def_ControlsPushColor = &HC08080
Const m_def_AutoSize = True
Const m_def_CloseButton = True
Const m_def_MinimizeButton = True
Const m_def_MaximizeButton = True
Const m_def_TitleBarFocusColor = &HE4C0C0
Const m_def_TitleBarNoFocusColor = &H8000000B
Const m_def_FormBorderColor = 0

'Property Variables:
Dim m_ButtonDisableColor As OLE_COLOR
Dim m_TitleBarLayout As TBPos
Dim m_Resizable As Boolean
Dim m_ControlsBorderColor As OLE_COLOR
Dim m_ControlsBackColor As OLE_COLOR
Dim m_ControlsHoverColor As OLE_COLOR
Dim m_ControlsPushColor As OLE_COLOR
Dim m_AutoSize As Boolean
Dim m_CloseButton As Boolean
Dim m_MinimizeButton As Boolean
Dim m_MaximizeButton As Boolean
'Dim m_TitleBarIcon As Picture
Dim m_TitleBarFocusColor As OLE_COLOR
Dim m_TitleBarNoFocusColor As OLE_COLOR
Dim m_FormBorderColor As OLE_COLOR

'Event Declarations:
'Event HasFocus()
'Event NoFocus()
Event CloseButtonClicked()
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    DrawWindow
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleBarCaption,TitleBarCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = TitleBarCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TitleBarCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleBarCaption,TitleBarCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = TitleBarCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TitleBarCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    DrawWindow
    UserControl.Refresh
End Sub






Private Sub Image1_Click()

RaiseEvent CloseButtonClicked

Unload UserControl.Parent

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture1.BackColor = m_ControlsPushColor

End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture1.BackColor = m_ControlsBackColor

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture2.BackColor = m_ControlsPushColor

End Sub


Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture2.BackColor = m_ControlsBackColor

If UserControl.Parent.WindowState = 0 Then
    UserControl.Parent.WindowState = vbMaximized
    SetNewSize
Else
    UserControl.Parent.WindowState = 0
    SetNewSize
End If

End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture3.BackColor = m_ControlsPushColor

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture3.BackColor = m_ControlsBackColor

UserControl.Parent.WindowState = vbMinimized

End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_Resizable = True Then
    Image4.MousePointer = 9
    If Button = 1 Then
        If UserControl.Width >= 1050 Then
            GetCursorPos MMove
            With UserControl.Parent
                .Width = MMove.x * Screen.TwipsPerPixelX - .Left
            End With
            Image4.Left = UserControl.Parent.Width - Image4.Width
        End If
        DrawWindow
    End If
End If

End Sub


Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Width < 1050 Then
    UserControl.Parent.Width = 1100
    DrawWindow
End If

End Sub


Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_Resizable = True Then
    Image5.MousePointer = 8
    If Button = 1 Then
        If UserControl.Width >= 1050 Then
            GetCursorPos MMove
            With UserControl.Parent
                .Width = MMove.x * Screen.TwipsPerPixelX - .Left
            End With
            'Image4.Left = UserControl.Parent.Width - Image4.Width
        End If
        If UserControl.Height >= 1050 Then
            GetCursorPos MMove
            With UserControl.Parent
                .Height = MMove.y * Screen.TwipsPerPixelY - .Top
            End With
            Image5.Left = UserControl.Parent.Width - Image5.Width
        End If
        DrawWindow
    End If
End If

End Sub


Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Width < 1050 Then
    UserControl.Parent.Width = 1100
    DrawWindow
End If

If UserControl.Height < 1050 Then
    UserControl.Parent.Height = 1100
    DrawWindow
End If

End Sub


Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_Resizable = True Then
    Image6.MousePointer = 7
    If Button = 1 Then
        If UserControl.Height >= 1050 Then
            GetCursorPos MMove
            With UserControl.Parent
                .Height = MMove.y * Screen.TwipsPerPixelY - .Top
            End With
            Image6.Width = UserControl.Parent.Width - Image5.Width
        End If
        DrawWindow
    End If
End If

End Sub


Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Height < 1050 Then
    UserControl.Parent.Height = 1100
    DrawWindow
End If

End Sub






Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture1.BackColor = m_ControlsPushColor

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_CloseButton = True Then
    If Picture1.BackColor <> m_ControlsBackColor Then
        Picture1.BackColor = m_ControlsBackColor
        ButBorder(0).Visible = False
        ButBorder(1).Visible = False
        ButBorder(2).Visible = False
        ButBorder(3).Visible = False
    End If
End If

If m_MaximizeButton = True Then
    If Picture2.BackColor <> m_ControlsBackColor Then
        Picture2.BackColor = m_ControlsBackColor
        ButBorder(4).Visible = False
        ButBorder(5).Visible = False
        ButBorder(6).Visible = False
        ButBorder(7).Visible = False
    End If
End If

If m_MinimizeButton = True Then
    If Picture3.BackColor <> m_ControlsBackColor Then
        Picture3.BackColor = m_ControlsBackColor
        ButBorder(8).Visible = False
        ButBorder(9).Visible = False
        ButBorder(10).Visible = False
        ButBorder(11).Visible = False
    End If
End If

If m_CloseButton = True Then
    Picture1.BackColor = m_ControlsHoverColor

    ButBorder(0).Visible = True
    ButBorder(1).Visible = True
    ButBorder(2).Visible = True
    ButBorder(3).Visible = True
End If

End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture1.BackColor = m_ControlsBackColor

End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture2.BackColor = m_ControlsPushColor

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_CloseButton = True Then
    If Picture1.BackColor <> m_ControlsBackColor Then
        Picture1.BackColor = m_ControlsBackColor
        ButBorder(0).Visible = False
        ButBorder(1).Visible = False
        ButBorder(2).Visible = False
        ButBorder(3).Visible = False
    End If
End If

If m_MaximizeButton = True Then
    If Picture2.BackColor <> m_ControlsBackColor Then
        Picture2.BackColor = m_ControlsBackColor
        ButBorder(4).Visible = False
        ButBorder(5).Visible = False
        ButBorder(6).Visible = False
        ButBorder(7).Visible = False
    End If
End If

If m_MinimizeButton = True Then
    If Picture3.BackColor <> m_ControlsBackColor Then
        Picture3.BackColor = m_ControlsBackColor
        ButBorder(8).Visible = False
        ButBorder(9).Visible = False
        ButBorder(10).Visible = False
        ButBorder(11).Visible = False
    End If
End If

Picture2.BackColor = m_ControlsHoverColor

ButBorder(4).Visible = True
ButBorder(5).Visible = True
ButBorder(6).Visible = True
ButBorder(7).Visible = True

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture2.BackColor = m_ControlsBackColor

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture3.BackColor = m_ControlsPushColor

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If m_CloseButton = True Then
    If Picture1.BackColor <> m_ControlsBackColor Then
        Picture1.BackColor = m_ControlsBackColor
        ButBorder(0).Visible = False
        ButBorder(1).Visible = False
        ButBorder(2).Visible = False
        ButBorder(3).Visible = False
    End If
End If

If m_MaximizeButton = True Then
    If Picture2.BackColor <> m_ControlsBackColor Then
        Picture2.BackColor = m_ControlsBackColor
        ButBorder(4).Visible = False
        ButBorder(5).Visible = False
        ButBorder(6).Visible = False
        ButBorder(7).Visible = False
    End If
End If

If m_MinimizeButton = True Then
    If Picture3.BackColor <> m_ControlsBackColor Then
        Picture3.BackColor = m_ControlsBackColor
        ButBorder(8).Visible = False
        ButBorder(9).Visible = False
        ButBorder(10).Visible = False
        ButBorder(11).Visible = False
    End If
End If

Picture3.BackColor = m_ControlsHoverColor

ButBorder(8).Visible = True
ButBorder(9).Visible = True
ButBorder(10).Visible = True
ButBorder(11).Visible = True

End Sub


Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture3.BackColor = &HDEB1B1

End Sub

Private Sub pSniff_NewMessage(ByVal hWnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)

If uMsg = WM_ACTIVATE Then
    Select Case wParam
        Case 0
            NoFocus
        Case Is > 0
            HasFocus
    End Select
End If

End Sub



Private Sub TitleBar_DblClick()

If m_MaximizeButton = True Then
    Picture2.BackColor = m_ControlsBackColor
End If

If m_MaximizeButton = True Then
    If UserControl.Parent.WindowState = 0 Then
        UserControl.Parent.WindowState = vbMaximized
        SetNewSize
    Else
        UserControl.Parent.WindowState = 0
        SetNewSize
    End If
End If

End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Parent.WindowState <> vbMaximized Then
    ism = True
    X1 = x
    Y1 = y
End If

End Sub


Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim I As Integer

If UserControl.Parent.WindowState <> vbMaximized Then
    If ism = True Then
        I = GetCursorPos(Pos)
        X2 = Pos.x * Screen.TwipsPerPixelX
        Y2 = Pos.y * Screen.TwipsPerPixelY
        UserControl.Parent.Move (X2 - X1), (Y2 - Y1)
    End If
End If

If HasGotFocus = True Then
    If m_CloseButton = True Then
        If Picture1.BackColor <> m_ControlsBackColor Then
            Picture1.BackColor = m_ControlsBackColor
            ButBorder(0).Visible = False
            ButBorder(1).Visible = False
            ButBorder(2).Visible = False
            ButBorder(3).Visible = False
        End If
    End If

    If m_MaximizeButton = True Then
        If Picture2.BackColor <> m_ControlsBackColor Then
            Picture2.BackColor = m_ControlsBackColor
            ButBorder(4).Visible = False
            ButBorder(5).Visible = False
            ButBorder(6).Visible = False
            ButBorder(7).Visible = False
        End If
    End If

    If m_MinimizeButton = True Then
        If Picture3.BackColor <> m_ControlsBackColor Then
            Picture3.BackColor = m_ControlsBackColor
            ButBorder(8).Visible = False
            ButBorder(9).Visible = False
            ButBorder(10).Visible = False
            ButBorder(11).Visible = False
        End If
    End If
End If
End Sub


Private Sub TitleBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Parent.WindowState <> vbMaximized Then
    If UserControl.Parent.Top < 0 Then UserControl.Parent.Top = 0
    If UserControl.Parent.Left < 0 Then UserControl.Parent.Left = 0
    If UserControl.Parent.Top > (Screen.Height - (UserControl.Parent.Height / 10)) Then UserControl.Parent.Top = Screen.Height * 9 / 10
    If UserControl.Parent.Left > (Screen.Width - (UserControl.Parent.Width / 10)) Then UserControl.Parent.Left = Screen.Width * 9 / 10
    ism = False
End If

End Sub

Private Sub TitleBarCaption_DblClick()

If m_MaximizeButton = True Then
    Picture2.BackColor = m_ControlsBackColor
End If

If m_MaximizeButton = True Then
    If UserControl.Parent.WindowState = 0 Then
        UserControl.Parent.WindowState = vbMaximized
        SetNewSize
    Else
        UserControl.Parent.WindowState = 0
        SetNewSize
    End If
End If

End Sub

Private Sub TitleBarCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ism = True
X1 = x
Y1 = y

End Sub

Private Sub TitleBarCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim I As Integer

If ism = True Then
    I = GetCursorPos(Pos)
    X2 = Pos.x * Screen.TwipsPerPixelX
    Y2 = Pos.y * Screen.TwipsPerPixelY
    UserControl.Parent.Move (X2 - X1), (Y2 - Y1)
End If

End Sub


Private Sub TitleBarCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If UserControl.Parent.Top < 0 Then UserControl.Parent.Top = 0
If UserControl.Parent.Left < 0 Then UserControl.Parent.Left = 0
If UserControl.Parent.Top > (Screen.Height - (UserControl.Parent.Height / 10)) Then UserControl.Parent.Top = Screen.Height * 9 / 10
If UserControl.Parent.Left > (Screen.Width - (UserControl.Parent.Width / 10)) Then UserControl.Parent.Left = Screen.Width * 9 / 10
ism = False

End Sub


Private Sub UserControl_Click()
'***************************************
'* Raise the Click Event               *
'***************************************
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()

'***************************************
'* Raise Double Click Event            *
'***************************************
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'***************************************
'* Check to see if control has focus   *
'***************************************
If HasGotFocus = True Then
    '***************************************
    '* Set Minimize, Maximize and Close    *
    '* Buttons Properties                  *
    '***************************************
    If m_CloseButton = True Then
        If Picture1.BackColor <> m_ControlsBackColor Then
            Picture1.BackColor = m_ControlsBackColor
            ButBorder(0).Visible = False
            ButBorder(1).Visible = False
            ButBorder(2).Visible = False
            ButBorder(3).Visible = False
        End If
    End If
    
    If m_MaximizeButton = True Then
        If Picture2.BackColor <> m_ControlsBackColor Then
            Picture2.BackColor = m_ControlsBackColor
            ButBorder(4).Visible = False
            ButBorder(5).Visible = False
            ButBorder(6).Visible = False
            ButBorder(7).Visible = False
        End If
    End If
    
    If m_MinimizeButton = True Then
        If Picture3.BackColor <> m_ControlsBackColor Then
            Picture3.BackColor = m_ControlsBackColor
            ButBorder(8).Visible = False
            ButBorder(9).Visible = False
            ButBorder(10).Visible = False
            ButBorder(11).Visible = False
        End If
    End If
End If

RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get TitleBarLayout() As TBPos
'    TitleBarLayout = m_TitleBarLayout
'End Property
'
'Public Property Let TitleBarLayout(ByVal New_TitleBarLayout As TBPos)
'    m_TitleBarLayout = New_TitleBarLayout
'    PropertyChanged "TitleBarLayout"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,0
''Public Property Get TitleBarThickness() As Variant
''    TitleBarThickness = m_TitleBarThickness
''End Property
''
''Public Property Let TitleBarThickness(ByVal New_TitleBarThickness As Variant)
''    m_TitleBarThickness = New_TitleBarThickness
''    PropertyChanged "TitleBarThickness"
''End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=14,0,0,0
'''Public Property Get TitleBarBorder() As Boolean
'''    TitleBarBorder = m_TitleBarBorder
'''End Property
'''
'''Public Property Let TitleBarBorder(ByVal New_TitleBarBorder As Boolean)
'''    m_TitleBarBorder = New_TitleBarBorder
'''    PropertyChanged "TitleBarBorder"
'''End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,0
'Public Property Get TitleBarBorderColor() As OLE_COLOR
'    TitleBarBorderColor = m_TitleBarBorderColor
'End Property
'
'Public Property Let TitleBarBorderColor(ByVal New_TitleBarBorderColor As OLE_COLOR)
'    m_TitleBarBorderColor = New_TitleBarBorderColor
'    PropertyChanged "TitleBarBorderColor"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,0
''Public Property Get FormBorder() As Boolean
''    FormBorder = m_FormBorder
''End Property
''
''Public Property Let FormBorder(ByVal New_FormBorder As Variant)
''    m_FormBorder = New_FormBorder
''    PropertyChanged "FormBorder"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,0
''Public Property Get FormBorderStyle() As BDStyle
''    FormBorderStyle = m_FormBorderStyle
''End Property
''
''Public Property Let FormBorderStyle(ByVal New_FormBorderStyle As BDStyle)
''    m_FormBorderStyle = New_FormBorderStyle
''    PropertyChanged "FormBorderStyle"
''End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TitleBarFocusColor() As OLE_COLOR
Attribute TitleBarFocusColor.VB_Description = "Color of the title bar when the form gets focus"
    TitleBarFocusColor = m_TitleBarFocusColor
    DrawWindow
End Property

Public Property Let TitleBarFocusColor(ByVal New_TitleBarFocusColor As OLE_COLOR)
    m_TitleBarFocusColor = New_TitleBarFocusColor
    DrawWindow
    PropertyChanged "TitleBarFocusColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TitleBarNoFocusColor() As OLE_COLOR
Attribute TitleBarNoFocusColor.VB_Description = "The color of the title bar when the form losses focus"
    TitleBarNoFocusColor = m_TitleBarNoFocusColor
End Property

Public Property Let TitleBarNoFocusColor(ByVal New_TitleBarNoFocusColor As OLE_COLOR)
    m_TitleBarNoFocusColor = New_TitleBarNoFocusColor
    PropertyChanged "TitleBarNoFocusColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get ShowTitleBar() As Variant
'    ShowTitleBar = m_ShowTitleBar
'End Property
'
'Public Property Let ShowTitleBar(ByVal New_ShowTitleBar As Variant)
'    m_ShowTitleBar = New_ShowTitleBar
'    PropertyChanged "ShowTitleBar"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=14,0,0,0
''Public Property Get ParentForm() As Form
''    ParentForm = m_ParentForm
''End Property
''
''Public Property Let ParentForm(ByVal New_ParentForm As Form)
''    m_ParentForm = New_ParentForm
''    PropertyChanged "ParentForm"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleBarCaption,TitleBarCaption,-1,Caption
Public Property Get TitleCaption() As String
Attribute TitleCaption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    TitleCaption = TitleBarCaption.Caption
    DrawWindow
End Property

Public Property Let TitleCaption(ByVal New_TitleCaption As String)
    TitleBarCaption.Caption() = New_TitleCaption
    UserControl.Parent.Caption = New_TitleCaption
    PropertyChanged "TitleCaption"
    DrawWindow
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FormBorderColor() As OLE_COLOR
Attribute FormBorderColor.VB_Description = "What color is the border"
    FormBorderColor = m_FormBorderColor
    DrawWindow
End Property

Public Property Let FormBorderColor(ByVal New_FormBorderColor As OLE_COLOR)
    m_FormBorderColor = New_FormBorderColor
    PropertyChanged "FormBorderColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
Debug.Print "init"
    m_TitleBarFocusColor = m_def_TitleBarFocusColor
    m_TitleBarNoFocusColor = m_def_TitleBarNoFocusColor
    m_FormBorderColor = m_def_FormBorderColor
    m_MinimizeButton = m_def_MinimizeButton
    m_MaximizeButton = m_def_MaximizeButton
    
    '***************************************
    '* Set UserControl to the size of      *
    '* the parent form                     *
    '***************************************
    With UserControl
        .Width = .Parent.Width
        .Height = .Parent.Height
    End With
    
    m_AutoSize = m_def_AutoSize
    m_CloseButton = m_def_CloseButton
    m_TitleBarLayout = m_def_TitleBarLayout
    m_Resizable = m_def_Resizable
    m_ControlsBorderColor = m_def_ControlsBorderColor
    m_ControlsBackColor = m_def_ControlsBackColor
    m_ControlsHoverColor = m_def_ControlsHoverColor
    m_ControlsPushColor = m_def_ControlsPushColor
    
    DrawWindow
    
    m_ButtonDisableColor = m_def_ButtonDisableColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Resizable = PropBag.ReadProperty("Resizable", m_def_Resizable)
    m_MaximizeButton = PropBag.ReadProperty("MaximizeButton", m_def_MaximizeButton)
    
    If m_Resizable = False Then m_MaximizeButton = False
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000B)
    TitleBarCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set TitleBarCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TitleBarFocusColor = PropBag.ReadProperty("TitleBarFocusColor", m_def_TitleBarFocusColor)
    m_TitleBarNoFocusColor = PropBag.ReadProperty("TitleBarNoFocusColor", m_def_TitleBarNoFocusColor)
    TitleBarCaption.Caption = PropBag.ReadProperty("TitleCaption", "Caption")
    m_FormBorderColor = PropBag.ReadProperty("FormBorderColor", m_def_FormBorderColor)
    m_MinimizeButton = PropBag.ReadProperty("MinimizeButton", m_def_MinimizeButton)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_CloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
    m_TitleBarLayout = PropBag.ReadProperty("TitleBarLayout", m_def_TitleBarLayout)
    Set TitleBarCaption.Font = PropBag.ReadProperty("TitleCaptionFont", Ambient.Font)
    TitleBarCaption.ForeColor = PropBag.ReadProperty("TitleCaptionColor", &H80000012)
    m_ControlsBorderColor = PropBag.ReadProperty("ControlsBorderColor", m_def_ControlsBorderColor)
    m_ControlsBackColor = PropBag.ReadProperty("ControlsBackColor", m_def_ControlsBackColor)
    m_ControlsHoverColor = PropBag.ReadProperty("ControlsHoverColor", m_def_ControlsHoverColor)
    m_ControlsPushColor = PropBag.ReadProperty("ControlsPushColor", m_def_ControlsPushColor)
    Set Picture = PropBag.ReadProperty("TitleBarIcon", Nothing)
    
    '***************************************
    '* Redraw the Control                  *
    '***************************************
If Ambient.UserMode = True Then
    Set pSniff = New SmartSubClass
    pSniff.SubClassHwnd UserControl.Parent.hWnd, True
End If
    DrawWindow
    
    m_ButtonDisableColor = PropBag.ReadProperty("ButtonDisableColor", m_def_ButtonDisableColor)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000B)
    Call PropBag.WriteProperty("ForeColor", TitleBarCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", TitleBarCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleBarFocusColor", m_TitleBarFocusColor, m_def_TitleBarFocusColor)
    Call PropBag.WriteProperty("TitleBarNoFocusColor", m_TitleBarNoFocusColor, m_def_TitleBarNoFocusColor)
    Call PropBag.WriteProperty("TitleCaption", TitleBarCaption.Caption, "Caption")
    Call PropBag.WriteProperty("FormBorderColor", m_FormBorderColor, m_def_FormBorderColor)
    Call PropBag.WriteProperty("MinimizeButton", m_MinimizeButton, m_def_MinimizeButton)
    Call PropBag.WriteProperty("MaximizeButton", m_MaximizeButton, m_def_MaximizeButton)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
    Call PropBag.WriteProperty("TitleBarLayout", m_TitleBarLayout, m_def_TitleBarLayout)
    Call PropBag.WriteProperty("TitleCaptionFont", TitleBarCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleCaptionColor", TitleBarCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Resizable", m_Resizable, m_def_Resizable)
    Call PropBag.WriteProperty("ControlsBorderColor", m_ControlsBorderColor, m_def_ControlsBorderColor)
    Call PropBag.WriteProperty("ControlsBackColor", m_ControlsBackColor, m_def_ControlsBackColor)
    Call PropBag.WriteProperty("ControlsHoverColor", m_ControlsHoverColor, m_def_ControlsHoverColor)
    Call PropBag.WriteProperty("ControlsPushColor", m_ControlsPushColor, m_def_ControlsPushColor)
    Call PropBag.WriteProperty("TitleBarIcon", Picture, Nothing)
    
    '***************************************
    '* Redraw the window                   *
    '***************************************
    DrawWindow
    Call PropBag.WriteProperty("ButtonDisableColor", m_ButtonDisableColor, m_def_ButtonDisableColor)
End Sub

Private Sub UserControl_Resize()

'***************************************
'* Set UserControl to size of the      *
'* parent form                         *
'***************************************
With UserControl
    .Width = .Parent.Width
    .Height = .Parent.Height
End With

'***************************************
'* Redraw the window                   *
'***************************************
DrawWindow
    
'***************************************
'* Set the Min, Max, and Close Buttons *
'* Properties                          *
'***************************************
If Picture1.BackColor <> m_ControlsBackColor Then
    Picture1.BackColor = m_ControlsBackColor
    ButBorder(0).Visible = False
    ButBorder(1).Visible = False
    ButBorder(2).Visible = False
    ButBorder(3).Visible = False
End If

If Picture2.BackColor <> m_ControlsBackColor Then
    Picture2.BackColor = m_ControlsBackColor
    ButBorder(4).Visible = False
    ButBorder(5).Visible = False
    ButBorder(6).Visible = False
    ButBorder(7).Visible = False
End If

If Picture3.BackColor <> m_ControlsBackColor Then
    Picture3.BackColor = m_ControlsBackColor
    ButBorder(8).Visible = False
    ButBorder(9).Visible = False
    ButBorder(10).Visible = False
    ButBorder(11).Visible = False
End If

'***************************************
'* Refresh UserControl, May not be     *
'* needed                              *
'***************************************
UserControl.Refresh
    
RaiseEvent Resize
    
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MinimizeButton() As Boolean
Attribute MinimizeButton.VB_Description = "Sets wether or not Form will have a minimize button"
    MinimizeButton = m_MinimizeButton
    DrawWindow
End Property

Public Property Let MinimizeButton(ByVal New_MinimizeButton As Boolean)
    m_MinimizeButton = New_MinimizeButton
    PropertyChanged "MinimizeButton"
    DrawWindow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MaximizeButton() As Boolean
Attribute MaximizeButton.VB_Description = "Sets wether or not Form will have a maximize button"
    MaximizeButton = m_MaximizeButton
    DrawWindow
End Property

Public Property Let MaximizeButton(ByVal New_MaximizeButton As Boolean)
    m_MaximizeButton = New_MaximizeButton
    PropertyChanged "MaximizeButton"
    DrawWindow
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=11,0,0,0
'Public Property Get TitleBarIcon() As Picture
'    Set TitleBarIcon = m_TitleBarIcon
'End Property
'
'Public Property Set TitleBarIcon(ByVal New_TitleBarIcon As Picture)
'    Set m_TitleBarIcon = New_TitleBarIcon
'    PropertyChanged "TitleBarIcon"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=0,0,0,True
''Public Property Get ShowFormFocus() As Boolean
''    ShowFormFocus = m_ShowFormFocus
''End Property
''
''Public Property Let ShowFormFocus(ByVal New_ShowFormFocus As Boolean)
''    m_ShowFormFocus = New_ShowFormFocus
''    PropertyChanged "ShowFormFocus"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=0,0,0,True
''Public Property Get TitleBarFocusEffect() As Boolean
''    TitleBarFocusEffect = m_TitleBarFocusEffect
''End Property
''
''Public Property Let TitleBarFocusEffect(ByVal New_TitleBarFocusEffect As Boolean)
''    m_TitleBarFocusEffect = New_TitleBarFocusEffect
''    PropertyChanged "TitleBarFocusEffect"
''End Property
'''
''''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''''MemberInfo=0,0,0,True
'''Public Property Get ShowTitleBar() As Boolean
'''    ShowTitleBar = m_ShowTitleBar
'''End Property
'''
'''Public Property Let ShowTitleBar(ByVal New_ShowTitleBar As Boolean)
'''    m_ShowTitleBar = New_ShowTitleBar
'''    PropertyChanged "ShowTitleBar"
'''    DrawTitle
'''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CloseButton() As Boolean
Attribute CloseButton.VB_Description = "Sets wether or not the form will have a Close Button"
    CloseButton = m_CloseButton
    DrawWindow
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    m_CloseButton = New_CloseButton
    PropertyChanged "CloseButton"
    DrawWindow
End Property


Private Sub DrawTitle()

Dim I As Integer

With UserControl
    .Width = UserControl.Parent.Width
    .Height = UserControl.Parent.Height
End With

'***************************************
'* Draw the TitleBar                   *
'***************************************
With TitleBar
    .Visible = True
    .BackColor = m_TitleBarFocusColor
    .Top = 0
    .Left = 0
    .Width = UserControl.Parent.Width
    .Height = 300
End With


'***************************************
'* Set the borders of the Min, Max,    *
'* and Close buttons to thier proper   *
'* colors and visibility               *
'***************************************
For I = 0 To 11
    ButBorder(I).BorderColor = m_ControlsBorderColor
    ButBorder(I).Visible = False
Next I

'***************************************
'* Clip the caption if it starts       *
'* to run behind the Min, Max, and     *
'* Close Buttons                       *
'***************************************
If TitleBarCaption.Width > Picture3.Left - 90 Then
    TitleBarCaption.AutoSize = False
    TitleBarCaption.Width = Picture3.Left - TitleBarCaption.Left
Else
    TitleBarCaption.AutoSize = True
    TitleBarCaption.Refresh
End If

'***************************************
'* Draw the TitleBar 3D lines if       *
'* the caption isn't taking up all     *
'* the space                           *
'***************************************
If (Picture3.Left - (TitleBarCaption.Left + TitleBarCaption.Width)) > 15 Then
    For I = 0 To 2
        TitleDetail(I).Visible = True
        TitleDetailHi(I).Visible = True
        TitleDetail(I).X2 = Picture3.Left - 90
        TitleDetailHi(I).X2 = Picture3.Left - 90
        TitleDetail(I).X1 = TitleBarCaption.Left + TitleBarCaption.Width + 90
        TitleDetailHi(I).X1 = TitleBarCaption.Left + TitleBarCaption.Width + 90
    Next I
Else
    For I = 0 To 2
        TitleDetail(I).Visible = False
        TitleDetailHi(I).Visible = False
    Next I
End If

'***************************************
'* Set Min, Max, and Close Buttons     *
'* Properties                          *
'***************************************
If m_CloseButton = True Then
    Picture1.BackColor = m_ControlsBackColor
Else
    Picture1.BackColor = m_ButtonDisableColor
End If

If m_MaximizeButton = True Then
    Picture2.BackColor = m_ControlsBackColor
Else
    Picture2.BackColor = m_ButtonDisableColor
End If

If m_MinimizeButton = True Then
    Picture3.BackColor = m_ControlsBackColor
Else
    Picture3.BackColor = m_ButtonDisableColor
End If

Picture1.Enabled = m_CloseButton
Picture2.Enabled = m_MaximizeButton
Picture3.Enabled = m_MinimizeButton

Picture1.Left = TitleBar.Width - 60 - Picture1.Width
Picture2.Left = Picture1.Left - 35 - Picture2.Width
Picture3.Left = Picture2.Left - 25 - Picture3.Width

End Sub

Private Sub DrawBorder()

'***************************************
'* Draw the border on the form         *
'***************************************
'***************************************
'* Border1 = Right                     *
'* Border0 = Bottom                    *
'* Border2 = Left                      *
'***************************************
With Border(0)
    .BorderWidth = 1
    .BorderColor = m_FormBorderColor
    .X1 = 0
    .X2 = UserControl.Width
    .Y1 = UserControl.Height - 15
    .Y2 = UserControl.Height - 15
    .Refresh
End With

With Border(1)
    .BorderWidth = 1
    .BorderColor = m_FormBorderColor
    .X1 = UserControl.Width - 15
    .X2 = UserControl.Width - 15
    .Y1 = 300
    .Y2 = UserControl.Height
    .Refresh
End With

With Border(2)
    .BorderWidth = 1
    .BorderColor = m_FormBorderColor
    .X1 = 0
    .X2 = 0
    .Y1 = 300
    .Y2 = UserControl.Height
    .Refresh
End With


With Image4
    .Top = 300
    .Width = 15
    .Height = UserControl.Height - 300 - Image5.Height
    .Left = UserControl.Width - Image4.Width
End With

With Image5
    .Left = UserControl.Width - Image5.Width
    .Top = UserControl.Height - Image5.Height
End With

With Image6
    .Height = 15
    .Top = UserControl.Height - Image6.Height
    .Left = 0
    .Width = UserControl.Width - Image5.Width
End With




End Sub

Public Sub SetNewSize()

With UserControl
    .Width = .Parent.Width
    .Height = .Parent.Height
End With
    
DrawTitle

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,0
Public Property Get TitleBarLayout() As TBPos
Attribute TitleBarLayout.VB_Description = "Where is title bar displayed"
    TitleBarLayout = m_TitleBarLayout
    DrawWindow
End Property

Public Property Let TitleBarLayout(ByVal New_TitleBarLayout As TBPos)
    m_TitleBarLayout = New_TitleBarLayout
    DrawWindow
    PropertyChanged "TitleBarLayout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleBarCaption,TitleBarCaption,-1,Font
Public Property Get TitleCaptionFont() As Font
Attribute TitleCaptionFont.VB_Description = "Returns a Font object."
    Set TitleCaptionFont = TitleBarCaption.Font
End Property

Public Property Set TitleCaptionFont(ByVal New_TitleCaptionFont As Font)
    Set TitleBarCaption.Font = New_TitleCaptionFont
    PropertyChanged "TitleCaptionFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TitleBarCaption,TitleBarCaption,-1,ForeColor
Public Property Get TitleCaptionColor() As OLE_COLOR
Attribute TitleCaptionColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleCaptionColor = TitleBarCaption.ForeColor
    DrawWindow
End Property

Public Property Let TitleCaptionColor(ByVal New_TitleCaptionColor As OLE_COLOR)
    TitleBarCaption.ForeColor() = New_TitleCaptionColor
    DrawWindow
    PropertyChanged "TitleCaptionColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Resizable() As Boolean
    Resizable = m_Resizable
End Property

Public Property Let Resizable(ByVal New_Resizable As Boolean)
    m_Resizable = New_Resizable
    PropertyChanged "Resizable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ControlsBorderColor() As OLE_COLOR
    ControlsBorderColor = m_ControlsBorderColor
    DrawWindow
End Property

Public Property Let ControlsBorderColor(ByVal New_ControlsBorderColor As OLE_COLOR)
    m_ControlsBorderColor = New_ControlsBorderColor
    DrawWindow
    PropertyChanged "ControlsBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ControlsBackColor() As OLE_COLOR
    ControlsBackColor = m_ControlsBackColor
    DrawWindow
End Property

Public Property Let ControlsBackColor(ByVal New_ControlsBackColor As OLE_COLOR)
    m_ControlsBackColor = New_ControlsBackColor
    DrawWindow
    PropertyChanged "ControlsBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ControlsHoverColor() As OLE_COLOR
    ControlsHoverColor = m_ControlsHoverColor
    DrawWindow
End Property

Public Property Let ControlsHoverColor(ByVal New_ControlsHoverColor As OLE_COLOR)
    m_ControlsHoverColor = New_ControlsHoverColor
    DrawWindow
    PropertyChanged "ControlsHoverColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ControlsPushColor() As OLE_COLOR
    ControlsPushColor = m_ControlsPushColor
    DrawWindow
End Property

Public Property Let ControlsPushColor(ByVal New_ControlsPushColor As OLE_COLOR)
    m_ControlsPushColor = New_ControlsPushColor
    DrawWindow
    PropertyChanged "ControlsPushColor"
End Property


Private Sub DrawWindow()

If UserControl.Parent.WindowState = vbMinimized Then
Else
DrawTitle
DrawBorder
End If

End Sub


Public Sub HasFocus()

'***************************************
'* Reset when form has focus           *
'***************************************
HasGotFocus = True

Picture1.Enabled = True
Picture2.Enabled = True
Picture3.Enabled = True

DrawWindow

End Sub

Public Sub NoFocus()


'***************************************
'* Changes appearence of form once     *
'* it loses focus                      *
'***************************************
Dim I As Integer

TitleBar.BackColor = m_TitleBarNoFocusColor
Picture1.BackColor = m_TitleBarNoFocusColor
Picture2.BackColor = m_TitleBarNoFocusColor
Picture3.BackColor = m_TitleBarNoFocusColor

HasGotFocus = False
Picture1.Enabled = False
Picture2.Enabled = False
Picture3.Enabled = False

TitleBarCaption.ForeColor = vbBlack

For I = 0 To 2
    TitleDetail(I).Visible = False
    TitleDetailHi(I).Visible = False
Next I

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonDisableColor() As OLE_COLOR
    ButtonDisableColor = m_ButtonDisableColor
End Property

Public Property Let ButtonDisableColor(ByVal New_ButtonDisableColor As OLE_COLOR)
    m_ButtonDisableColor = New_ButtonDisableColor
    PropertyChanged "ButtonDisableColor"
End Property

