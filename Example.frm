VERSION 5.00
Object = "{D9FDB204-2D4F-4C34-864A-9D9289DB746F}#59.0#0"; "Menu.ocx"
Begin VB.Form Form1 
   Caption         =   "Custom Menu Example 2"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Enabled"
      Height          =   1815
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1815
      Begin VB.CheckBox Check9 
         Caption         =   "Use Left Image"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Animate"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Paste"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Copy"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Cut"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Exit"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "New"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Border && Selection Style"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1935
      Begin VB.ListBox List2 
         Height          =   645
         ItemData        =   "Example.frx":0000
         Left            =   120
         List            =   "Example.frx":000D
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Example.frx":0034
         Left            =   120
         List            =   "Example.frx":0041
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin Menu_Replacement.MenuCtl Menu1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemHotForeColor=   255
      BackColor       =   10454404
      ItemHotBackColor=   16761024
      OpenAnimated    =   -1  'True
      UseLeftImage    =   -1  'True
      LeftPicBackColor=   0
      MouseOverSelectionType=   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click + Icons"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mouse Down"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   1500
      Left            =   1920
      Picture         =   "Example.frx":0060
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":1362
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":16EC
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":1A76
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "By Daniel Taylor..."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":1E00
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":218A
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "Example.frx":2514
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Example for Menu.ocx Windows Menu Replacement Control
'Created By Daniel Taylor on 17 April 2001
'This code just demonstrates the use of my Menu.ocx file.
'it shows how to make a menu, have it show, and get the
'results when its clicked on and respond to these results.
'It just shows off a little of what you can do with the menu
'Also, the colors are all changable, try experimenting, and
'As always if you like the code, VOTE!!!!(Font is changeable, too)
'::::::To use it:::::
'first make a menu using the
'.AddItem ItemText, [ItemName], [Enabled], [Icon], [Forecolor], [HotForecolor], [HotBackcolor]
'then once all done, just show the menu using
'.ShowMenu [X], [Y]
'If X & Y are not specified, it opens where the cursor is at
'When the mouse is let up the menu closes and gives back the
'Index and Text of the Item clicked, and then i set it up
'to either display certain messages or to exit the program.
'Also, I used a few image controls instead of an imagelist,
'but i think it would be easier to use the imagelist for the icons...

Private Sub Check8_Click()
    'tell the menu to open animated if check8 is checked,
    'else, don't animate
    Menu1.OpenAnimated = Check8
End Sub

Private Sub Check9_Click()
    Menu1.UseLeftImage = Check9
End Sub

Private Sub Command1_Click()
    'create normal menu, no icons
    Menu1.AddItem "New...", , Check1, , RGB(100, 100, 0)
    Menu1.AddItem "Save...", , Check2
    Menu1.AddItem "Delete...", , Check3, , , RGB(150, 0, 150)
    Menu1.AddItem "", "Seperator"
    Menu1.AddItem "Cut", , Check5, , , , RGB(0, 255, 0)
    Menu1.AddItem "Copy", , Check6, , RGB(0, 100, 100)
    Menu1.AddItem "Paste", , Check7
    Menu1.AddItem "", "seperator"
    Menu1.AddItem "Exit", , Check4
    'show the menu
    Menu1.ShowMenu , , Image7.Picture
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'just make and show the menu again
    Call Command1_Click
End Sub

Private Sub Command3_click()
    'this time build it with icons...
    'set the icon usage to true
    'don't specify an icon, and the space next to the text
    'is blank, just like the windows menus... (i.e. the exit item)
    Menu1.UseIcons = True
    Menu1.AddItem "New...", , Check1, Image1.Picture, RGB(100, 100, 0)
    Menu1.AddItem "Save...", , Check2, Image2.Picture
    Menu1.AddItem "Delete...", , Check3, Image3.Picture, , RGB(150, 0, 150)
    Menu1.AddItem "", "Seperator"
    Menu1.AddItem "Cut", , Check5, Image4.Picture, , , RGB(0, 255, 0)
    Menu1.AddItem "Copy", , Check6, Image5.Picture, RGB(0, 100, 100)
    Menu1.AddItem "Paste", , Check7, Image6.Picture
    Menu1.AddItem "", "seperator"
    Menu1.AddItem "Exit", , Check4
    'once again show the menu
    Menu1.ShowMenu , , Image7.Picture
    'make sure to reset the icon usage
    Menu1.UseIcons = False
End Sub

Private Sub Form_Load()
    List1.ListIndex = 0
    List2.ListIndex = 1
End Sub

Private Sub List1_Click()
    'change the borderstyle of the menu
    Menu1.Style = List1.ListIndex
End Sub

Private Sub List2_Click()
    Menu1.MouseOverSelectionType = List2.ListIndex
End Sub

Private Sub Menu1_ItemClicked(Index As Integer, Text As String, Name As String)
    'get results of menu when it closes, then process the text
    If Text = "Exit" Then
        DoEvents
        Unload Me
    ElseIf Text = "New..." Then
        MsgBox "New Not Possible..."
    Else
        MsgBox "You clicked an item with the index of " & Index & " and the text: " & Text
    End If
End Sub
