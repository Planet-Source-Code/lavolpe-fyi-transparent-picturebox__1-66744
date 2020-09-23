VERSION 5.00
Begin VB.UserControl ucTransPicBox 
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   HitBehavior     =   0  'None
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   47
   Windowless      =   -1  'True
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "ucTransPicBox.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   660
   End
End
Attribute VB_Name = "ucTransPicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This was just a fun project based on a "What If" question...

' This usercontrol is transparent and windowless, so it should work in VB fine, but
' if compiled, may not work with other languages if their containers don't support
' transparent, windowless controls.

' By placing a transparent usercontrol between the form and another control, we can
' allow the usercontrol (UC) to capture the form's contents and then pass that off
' to the target control.  This is a non-subclassing approach.  In order for this to
' work, you simply need to ensure 3 things:

' 1. The UC is placed exactly under the target control where you want the form
'   contents captured. If a control has an hWnd property and is under the target
'   control, that will not be captured. Only windowless controls and the form itself
'   is captured. Windowless controls? shapes, labels, image controls, etc.
'   See the Initialize routine which assists in placing the UC
' 2. The UC must be top most in the ZOrder so that it sits over all other windowless
'   controls. This UC is also windowless and making it top most does not affect
'   windowed controls (commandbuttons, textboxes, etc).
'   Simply call UserControl1.ZOrder to set top most
' 3. The UC must be associated with a control that has a Refresh method. Most common
'   associated controls you might use this with:  PictureBox, ImageControl.
'   And both of those do have .Refresh method.



' P.S. You can change the image within this UC's Image1 object to anything you want.
' The image is only visible during design time.

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

' used to create a stdPicture from a bitmap/icon handle
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private u_Parent As Control

Public Sub Initialize(ByVal FormHwnd As Long, ByVal TargetContHwnd As Long, _
                    ByRef TargetX As Long, ByRef TargetY As Long, _
                    TargetRef As Control)
        
    ' routine has 2 purposes:
    ' 1. Associate a visible control with this control.
    '   The control must have a Refresh method to work properly (i.e., PictureBox, ImageControl, etc)
    ' 2. Help you position the usercontrol so it is aligned under the object you
    '   want to treat as "transparent"
    
    
    On Error Resume Next
    If FormHwnd = 0 Then Exit Sub
    If TargetRef Is Nothing Then Exit Sub
    
    Set u_Parent = TargetRef
    
    Dim cRect As RECT, wRect As RECT, tPT As POINTAPI
    
    If TargetContHwnd = 0 Then TargetContHwnd = FormHwnd ' use FormHwnd in this case
        
    ' get the control's client area (no borders) this usercontrol will be placed under
    GetClientRect TargetContHwnd, cRect
    ' convert client coords to screen coords
    tPT.X = cRect.Left
    tPT.Y = cRect.Top
    ClientToScreen TargetContHwnd, tPT
    ' convert screen coords to the form's client coords
    ScreenToClient FormHwnd, tPT
    
    ' send back the X,Y coordinates. This will be the usercontrol's top,left
    ' position (in pixels). So convert pixels to your form's ScaleMode when
    ' the routine returns
    TargetX = tPT.X + TargetX
    TargetY = tPT.Y + TargetY
    
    ' once this routine returns, use the TargetX,TargetY coords to position the
    ' usercontrol and don't forget to resize it to the proper width/height that
    ' you need.  Use the UserControl1.Move method

End Sub

Public Property Get Image() As StdPicture
    ' use this to set your imagecontrol, picturebox, etc, picture property
    Set Image = Image1
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
'   When this UC is associated with ImageControls, calling refresh from this UC
'   doesn't always cause the ImageControl to repaint when the form's background color
'   changes. This doesn't seem to apply when associated with PictureBoxes.
'   A minor bug, so this is the workaround...
    If Ambient.UserMode = True Then ' in runtime
        Image1.Visible = True
        Image1.Visible = False
    End If
End Sub

Private Sub UserControl_Paint()
    If Ambient.UserMode = False Then Exit Sub   ' n/a if in design mode
    If Not u_Parent Is Nothing Then
        On Error Resume Next
        Dim tDC As Long, tOldBmp As Long
        tDC = CreateCompatibleDC(UserControl.hdc)
        tOldBmp = SelectObject(tDC, Image1.Picture.handle)
        ' copy form dc contents to our image control
        BitBlt tDC, 0, 0, ScaleWidth, ScaleHeight, UserControl.hdc, 0, 0, vbSrcCopy
        SelectObject tDC, tOldBmp
        ' tell parent to refresh
        u_Parent.Refresh
                                Debug.Print "... painting "; Timer
    End If
End Sub

Private Sub CreateStdPicture()
    ' function creates a stdPicture object from a image handle
    Dim lpPictDesc As PictDesc, aGUID(0 To 3) As Long
    Dim dDC As Long, tPic As IPictureDisp
    
    dDC = GetDC(0&)
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hHandle = CreateCompatibleBitmap(dDC, ScaleWidth, ScaleHeight)
        .hPal = 0
    End With
    ReleaseDC 0&, dDC
    
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Set Image1.Picture = Nothing
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, tPic)
    Set Image1.Picture = tPic
    
End Sub

Private Sub UserControl_Resize()
    If Ambient.UserMode = True Then ' in runtime mode
        CreateStdPicture            ' resize our image control
    Else
        UserControl.Width = ScaleX(Image1.Width, vbPixels, vbContainerSize)
        UserControl.Height = ScaleY(Image1.Height, vbPixels, vbContainerSize)
    End If
End Sub


Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    ' when in design mode, this allows user to click on & move a transparent control
    If Ambient.UserMode = False Then HitResult = vbHitResultHit
End Sub


Private Sub UserControl_Show()
    If Ambient.UserMode = True Then
        ' basically, make this UC invisible at run time
        Image1.Stretch = False
        Image1.Visible = False
    End If
End Sub
