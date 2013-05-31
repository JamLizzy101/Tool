VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Declare variables for panel instances
Private WithEvents PanelTop As cWidgetForm
Attribute PanelTop.VB_VarHelpID = -1
Private WithEvents PanelLeft As cWidgetForm
Attribute PanelLeft.VB_VarHelpID = -1
Private WithEvents PanelRight As cWidgetForm
Attribute PanelRight.VB_VarHelpID = -1
Private spFso As cFSO
Dim LeftBut() As cwButton
Dim SpellAddBtn As cwButton
Dim SpellFrmBox As cwTextBox, SpellFileName As cwTextBox
Dim Hedder As cwLabel
Dim ButCap() As String
Dim PanRTag As Integer
Dim PanLTag As Integer

Public WithEvents Form As cWidgetForm 'Declare Main Form instance
Attribute Form.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Dim i As Integer
    PanLTag = 0
    PanRTag = 0
    'Create Main Form without a WidgetRoot
    Set Form = Cairo.WidgetForms.Create(vbSizable, "Tool", False)
        Form.BackColor = RGB(0, 0, 0)
        
    'Create panel instances
    Set PanelTop = Cairo.WidgetForms.CreateChild(Form.hWnd)
        PanelTop.WidgetRoot.BackColor = RGB(255, 255, 255)
        PanelTop.WidgetRoot.ImageKey = "TopPanBg"
        
    Set PanelLeft = Cairo.WidgetForms.CreateChild(Form.hWnd)
        PanelLeft.WidgetRoot.BackColor = RGB(255, 255, 255)
        PanelLeft.WidgetRoot.ImageKey = "LeftPanBg"
        
        
        ReDim Preserve ButCap(0)
        ButCap(0) = "Spelling List"
        
        ReDim Preserve ButCap(1)
        ButCap(1) = "SVG Viewer"
        
        For i = 2 To 9
            ReDim Preserve ButCap(i)
            ButCap(i) = "Empty Button"
        Next
        
        ReDim LeftBut(9)
        For i = 0 To 9
            Set LeftBut(i) = PanelLeft.Widgets.Add(New cwButton, "PanLeftBut" & i, 5, 5 + i * 35, 160, 25, True)
            LeftBut(i).Caption = ButCap(i)
        Next
        
    Set PanelRight = Cairo.WidgetForms.CreateChild(Form.hWnd)
        PanelRight.WidgetRoot.BackColor = RGB(255, 255, 255)
        PanelRight.WidgetRoot.ImageKey = "RightPanBg"
End Sub

Private Sub spellLayout()
    PanelRight.Widgets.RemoveAll
    PanelRight.Visible = False
    Set Hedder = PanelRight.Widgets.Add(New cwLabel, "SpellHeader", 5, 5, PanelRight.Width - 10, 35, True)
        Hedder.Caption = "Spelling Lists"
        Hedder.Widget.FontName = "Century Gothic"
        Hedder.Widget.FontBold = True
        Hedder.Widget.FontSize = 18
        Hedder.Widget.ForeColor = RGB(160, 30, 5)
        Hedder.Widget.Alpha = 0
        
    Set SpellFrmBox = PanelRight.Widgets.Add(New cwTextBox, "RightSpellBox", 5, Hedder.Widget.Height + 10, PanelRight.Width - 10, PanelRight.Height - 150, True)
        SpellFrmBox.Widget.BackColor = RGB(255, 255, 255)
        SpellFrmBox.MultiLine = True
        
    
    Set SpellAddBtn = PanelRight.Widgets.Add(New cwButton, "SpellAddBtn", 5, SpellFrmBox.Widget.Top + SpellFrmBox.Widget.Height + 5, 140, 25, True)
        SpellAddBtn.Caption = "Add Spelling"
    
    PanelRight.Visible = True
    PanelRight.Refresh
    
End Sub

Private Sub svgLayout()
    PanelRight.Widgets.RemoveAll
    PanelRight.Visible = False
    Set Hedder = PanelRight.Widgets.Add(New cwLabel, "SvgHeader", 5, 5, PanelRight.Width - 10, 35, True)
        Hedder.Caption = "SVG Viewer"
        Hedder.Widget.FontName = "Century Gothic"
        Hedder.Widget.FontBold = True
        Hedder.Widget.FontSize = 18
        Hedder.Widget.ForeColor = RGB(160, 30, 5)
        Hedder.Widget.Alpha = 0
        PanelRight.Visible = True
        PanelRight.Refresh
End Sub

Private Sub Form_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
    Dim sKey As String, fKey As String
    
'    Debug.Print Sender.Widget.Key
'    Debug.Print "-------"
'    Debug.Print EventName
    Select Case EventName
        Case "W_Click"
            sKey = Sender.Widget.Key
            'Debug.Print Sender.Widget.Key
            Select Case sKey
                Case "PanLeftBut0"
                    spellLayout
                    'Made more sense that the widget tag remained in the BubblingEvent
                    PanelRight.Tag = PanelRight.Width
                    PanRTag = PanelRight.Tag
                Case "PanLeftBut1"
                    svgLayout
                Case "SpellAddBtn"
                    SpellAdd
            End Select
    End Select
    PanelLeft.Tag = PanelLeft.Width
    PanLTag = PanelLeft.Tag
End Sub

Private Sub SpellAdd()
    Set spFso = New_c.FSO
    
    spFso.CreateFileStream App.Path & "/test.txt", STRM_READWRITE, FA_NORMAL
    spFso.WriteTextContent App.Path & "/test.txt", SpellFrmBox.Text, True, True
    Debug.Print SpellFrmBox.Text
End Sub


Private Sub Form_ResizeWithDimensions(ByVal NewWidth As Long, ByVal NewHeight As Long)
    Const OffsPxl As Long = 5
    'Dim dxHalf As Long: dxHalf = (NewWidth - OffsPxl) \ 2  'Defines half width of form, but I don't need it
    Dim dyTop As Long: dyTop = (NewHeight - OffsPxl) \ 5    'Defined a Height - 1/5 of form height
    Dim dyLeft As Long: dyLeft = (NewWidth - OffsPxl) \ 6   'Defined a width - 1/6 of the form width
    
    
        
    'Dynamic form resizing
    PanelTop.Move OffsPxl, OffsPxl, NewWidth - 2 * OffsPxl, dyTop - OffsPxl
    PanelLeft.Move OffsPxl, dyTop + OffsPxl, dyLeft - OffsPxl, ((dyTop * 4) + 2) - OffsPxl  'Took dyTop val (1/5 height of form) * 4 to get remainder height
    PanelRight.Move dyLeft + OffsPxl, dyTop + OffsPxl, ((dyLeft * 5) + 2) - OffsPxl, PanelLeft.Height   'Took dyLeft val (1/6 width of form) * 5 to get remainder width
    
    
        
    If PanRTag <> 0 Then PanelRight.WidgetRoot.Zoom = (PanelRight.Width / PanRTag)
    If PanLTag <> 0 Then PanelLeft.WidgetRoot.Zoom = (PanelLeft.Width / PanLTag)
    Debug.Print PanelLeft.Width & " x " & PanelLeft.Height & " PanLTag: " & PanLTag
End Sub

Private Sub PanelTop_BubblingEvent(Sender As Object, EventName As String, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant)
Debug.Print EventName
End Sub

