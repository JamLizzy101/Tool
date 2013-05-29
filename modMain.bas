Attribute VB_Name = "modMain"
Option Explicit

Public New_c As New cConstructor, Cairo As cCairo
Public fMain As New cfMain

Sub Main()
    Set Cairo = New_c.Cairo
    
    LoadPicResources
    
    fMain.Form.Show
    Cairo.WidgetForms.EnterMessageLoop 'Is a requirement as there are no VB Forms
End Sub

Sub LoadPicResources()
    Cairo.ImageList.AddImage "TopPanBg", App.Path & "\gfx\topbg.jpg"
    Cairo.ImageList.AddImage "LeftPanBg", App.Path & "\gfx\leftbg.jpg"
    Cairo.ImageList.AddImage "RightPanBg", App.Path & "\gfx\rightbg.jpg"
End Sub
