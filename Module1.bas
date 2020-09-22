Attribute VB_Name = "Module1"
'Type and Public declaration module for Menu.ocx
'Created by Daniel Taylor on April 14, 2001
'This holds the data needed to make the menu

'type so each item can have its own properties
Public Type Item
    Name As String
    Text As String
    Enabled As Boolean
    Pic As StdPicture
    IForecolor As OLE_COLOR
    IHotForecolor As OLE_COLOR
    IHotBackcolor As OLE_COLOR
End Type

'all the items on the menu + other data holding variables
Public Items() As Item
Public ItemCount As Integer
Public HotItem As Integer
Public OldHotItem As Integer
Public MenuClosed As Boolean
Public LeftPic As StdPicture
''
Public m_Style As Style_Type
Public m_UseIcons As Boolean
Public m_BackColor As OLE_COLOR
Public m_OpenAnimated As Boolean
Public m_MenuAnimSpeed As Single
Public m_UseLeftImage As Boolean
Public m_LeftPicBackColor As OLE_COLOR
Public m_MouseOverSelectionType As Variant
