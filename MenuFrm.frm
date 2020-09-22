VERSION 5.00
Begin VB.Form MenuFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "MenuFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MenuFrm for Menu.ocx
'Created by Daniel Taylor on April 14, 2001
'This is the actual menu, but we need the usercontrol to
'access and show it through another program.
'its sort of messy because I didn't do it all at once,
'i kept adding stuff on as i went a long, some didn't work,
'other stuff did and now its sort of in a bunch of pieces...

Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim TPPX As Long, TPPY As Long
    
Private Sub Form_Load()
    HotItem = 0
    OldHotItem = HotItem
    TPPX = Screen.TwipsPerPixelX
    TPPY = Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > Me.Width Or Y < 0 Or Y > Me.Height Then
      If HotItem <> -1 Then
        OldHotItem = HotItem
        HotItem = -1
        DrawMenu False, False, False
      End If
    ElseIf CInt(((4 * TPPY) + Y) / TextHeight(Items(0).Text)) <> HotItem And X * TPPX > 0 And X < Me.Width Then
        OldHotItem = HotItem
        HotItem = CInt(((4 * TPPX) + Y) / TextHeight(Items(0).Text))
        DrawMenu False, False, False
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuClosed = True
    ReleaseCapture
    Unload Me
End Sub

Public Sub DrawMenu(Optional Drawborder As Boolean = True, Optional ResizeMenu As Boolean = True, Optional DrawIcons As Boolean = True, Optional OpenAnim As Boolean = False)
On Error Resume Next
Dim a As Integer, MaxTextWidth As Integer
Dim holdx As Single, holdy As Single
Dim z As Integer
If Drawborder = True Then
    Me.Cls
    If m_UseLeftImage = True Then
        Me.Line (2 * TPPX, 3 * TPPY)-(17 * TPPX, Me.Height - (4 * TPPY)), m_LeftPicBackColor, BF
        Me.PaintPicture LeftPic, 2 * TPPX, (Me.Height - (3 * TPPY)) - ScaleY(LeftPic.Height, vbHimetric, vbTwips)
    End If
End If
MaxTextWidth = 0
Me.CurrentY = 4 * TPPY
If ItemCount <> -1 Then
    If ResizeMenu = True Then
    For a = 0 To ItemCount
        If TextWidth(Items(a).Text) > MaxTextWidth Then
            MaxTextWidth = TextWidth(Items(a).Text)
        End If
    Next a
    If m_OpenAnimated = False Then
        Me.Height = (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
        If Me.Top + Me.Height > Screen.Height Then
            Me.Top = Screen.Height - Me.Height
        End If
        If m_UseIcons = False Then
            If m_UseLeftImage = False Then
                Me.Width = MaxTextWidth + (8 * TPPX)
            Else
                Me.Width = MaxTextWidth + (28 * TPPX)
            End If
        Else
            If m_UseLeftImage = False Then
                Me.Width = MaxTextWidth + (28 * TPPX)
            Else
                Me.Width = MaxTextWidth + (48 * TPPX)
            End If
        End If
        If Me.Left + Me.Width > Screen.Width Then
            Me.Left = Screen.Width - Me.Width
        End If
        If m_UseLeftImage = True Then
            holdx = Me.CurrentX
            holdy = Me.CurrentY
            Me.Line (2 * TPPX, 3 * TPPY)-(17 * TPPX, Me.Height - (4 * TPPY)), m_LeftPicBackColor, BF
            Me.PaintPicture LeftPic, 2 * TPPX, (Me.Height - (3 * TPPY)) - ScaleY(LeftPic.Height, vbHimetric, vbTwips)
            Me.CurrentX = holdx
            Me.CurrentY = holdy
        End If
    Else
      If OpenAnim = True Then
        Me.Height = 10
        Me.Width = 10
        For a = 0 To ItemCount
            If TextWidth(Items(a).Text) > MaxTextWidth Then
                MaxTextWidth = TextWidth(Items(a).Text)
            End If
        Next a
        For a = 0 To (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
            Me.Height = Me.Height + m_MenuAnimSpeed
            If Me.Height + Me.Top > Screen.Height Then
                Me.Top = Screen.Height - Me.Height
            End If
            If Me.Height > ((TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)) - m_MenuAnimSpeed Then
                Me.Height = (TextHeight(Items(0).Text) * (ItemCount + 1)) + (8 * TPPY)
                DrawMenu
                Exit For
            End If
            If m_UseIcons = False Then
                If m_UseLeftImage = False Then
                    z = 8
                Else
                    z = 28
                End If
            Else
                If m_UseLeftImage = False Then
                    z = 28
                Else
                    z = 64
                End If
            End If
            If Me.Width < (MaxTextWidth + (z * TPPX)) - m_MenuAnimSpeed Then
                Me.Width = Me.Width + m_MenuAnimSpeed
            Else
                Me.Width = MaxTextWidth + (z * TPPX)
            End If
            If Me.Width + Me.Left > Screen.Width Then
                Me.Left = Screen.Width - Me.Width
            End If
            DrawMenu , False
            DoEvents
        Next a
      End If
    End If
  End If
  For a = 0 To ItemCount
    If m_UseIcons = False Then
        If m_UseLeftImage = False Then
            z = 4
        Else
            z = 22
        End If
    Else
        If m_UseLeftImage = False Then
            z = 22
        Else
            z = 40
        End If
    End If
    Me.CurrentX = z * TPPX
    If Items(a).Enabled = True Then
        If Items(a).IForecolor <> Me.ForeColor Then
            Me.ForeColor = Items(a).IForecolor
        End If
        If a <> HotItem - 1 Then
            holdx = Me.CurrentX
            holdy = Me.CurrentY
            If m_UseIcons = True Then
                If DrawIcons = True Then
                    If m_UseLeftImage = False Then
                        Me.PaintPicture Items(a).Pic, TPPX * 4, holdy
                    Else
                        Me.PaintPicture Items(a).Pic, TPPX * 20, holdy
                    End If
                    Me.CurrentX = holdx
                    Me.CurrentY = holdy
                End If
            End If
            If Drawborder = False Then
                If a = OldHotItem - 1 Then
                    If Items(a).IHotBackcolor <> m_BackColor Then
                        Me.Line (holdx - TPPX, holdy)-(Me.Width - (5 * TPPX), holdy + TextHeight(Items(a).Text)), m_BackColor, BF
                        Me.CurrentX = holdx
                        Me.CurrentY = holdy
                    End If
                    Me.Print Items(a).Text
                Else
                    Me.CurrentY = Me.CurrentY + TextHeight(Items(a).Text)
                End If
            Else
                Me.Print Items(a).Text
            End If
        Else
            holdx = Me.CurrentX
            holdy = Me.CurrentY
            If m_UseIcons = True Then
                If DrawIcons = True Then
                    If m_UseLeftImage = False Then
                        Me.PaintPicture Items(a).Pic, TPPX * 4, holdy
                    Else
                        Me.PaintPicture Items(a).Pic, TPPX * 20, holdy
                    End If
                    Me.CurrentX = holdx
                    Me.CurrentY = holdy
                End If
            End If
            If Items(a).IHotBackcolor <> m_BackColor Then
                If m_MouseOverSelectionType = 0 Then
                    Me.Line (holdx - TPPX, holdy)-(Me.Width - (5 * TPPX), holdy + TextHeight(Items(a).Text)), Items(a).IHotBackcolor, BF
                Else
                    If m_MouseOverSelectionType = 2 Then
                        Me.DrawStyle = 2
                        Me.Line (holdx - TPPX, holdy)-(holdx + TextWidth(Items(a).Text), holdy + (TextHeight(Items(a).Text) - TPPY)), Items(a).IHotBackcolor, B
                        Me.DrawStyle = 0
                    Else
                        Me.Line (holdx - TPPX, holdy)-(holdx + TextWidth(Items(a).Text), holdy + (TextHeight(Items(a).Text) - TPPY)), Items(a).IHotBackcolor, B
                    End If
                End If
            End If
            Me.CurrentX = holdx
            Me.CurrentY = holdy
            Me.ForeColor = Items(a).IHotForecolor
            Me.Print Items(a).Text
            Me.ForeColor = Items(a).IForecolor
        End If
    Else
      Dim Color1 As OLE_COLOR, Color2 As OLE_COLOR
      Color1 = -1: Color2 = -1
      CheckForColors Me, Color1, Color2
      If LCase(Items(a).Name) <> "seperator" Then
        holdx = Me.CurrentX
        holdy = Me.CurrentY
        If m_UseIcons = True Then
            If DrawIcons = True Then
                If m_UseLeftImage = False Then
                    Me.PaintPicture Items(a).Pic, TPPX * 4, holdy
                Else
                    Me.PaintPicture Items(a).Pic, TPPX * 20, holdy
                End If
                Me.CurrentX = holdx
                Me.CurrentY = holdy
            End If
        End If
        Me.CurrentX = holdx + (1 * TPPX)
        Me.CurrentY = holdy + (1 * TPPY)
        Me.ForeColor = Color1
        Me.Print Items(a).Text
        Me.CurrentX = holdx
        Me.CurrentY = holdy
        Me.ForeColor = Color2
        Me.Print Items(a).Text
        Me.ForeColor = Items(a).IForecolor
      Else
        If m_UseLeftImage = False Then
            Me.CurrentX = 4 * TPPX
        Else
            Me.CurrentX = 20 * TPPX
        End If
        holdx = Me.CurrentX
        holdy = Me.CurrentY
        Me.Line (holdx, (holdy + (TextHeight(Items(a).Text) / 2) + 10))-(Me.Width - (5 * TPPY), (holdy + (TextHeight(Items(a).Text) / 2) + 10)), Color1
        Me.Line (holdx, holdy + (TextHeight(Items(a).Text) / 2))-(Me.Width - (5 * TPPY), holdy + (TextHeight(Items(a).Text) / 2)), Color2
        Me.CurrentX = holdx + TextHeight(Items(a).Text)
        Me.CurrentY = holdy + TextHeight(Items(a).Text)
      End If
    End If
  Next a
  If Drawborder = True Then
    If m_Style = Etch_Style Then
        Etch Me
    ElseIf m_Style = OutDent_Style Then
        OutLayered Me, 3
    Else
        PlainBorder Me
    End If
  End If
End If
End Sub
