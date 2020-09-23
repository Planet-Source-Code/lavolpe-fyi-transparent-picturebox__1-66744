VERSION 5.00
Begin VB.Form frmTPB 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmTPB.frx":0000
   ScaleHeight     =   3870
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Partialy ""Transparent"""
      Height          =   540
      Index           =   1
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      Width           =   1860
   End
   Begin VB.CheckBox Check1 
      Caption         =   "100% ""Transparent"""
      Height          =   540
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   1875
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Move Image"
      Height          =   495
      Left            =   5055
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "No Bubbles"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2535
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chg Lbl Text"
      Height          =   495
      Left            =   3705
      TabIndex        =   4
      Top             =   2535
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chg BkgColor"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2535
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2070
      Left            =   840
      ScaleHeight     =   2010
      ScaleWidth      =   5760
      TabIndex        =   1
      Top             =   1110
      Width           =   5820
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   480
         Top             =   390
         Visible         =   0   'False
         Width           =   4650
      End
   End
   Begin prjTPB.ucTransPicBox UserControl11 
      Height          =   750
      Left            =   6765
      TabIndex        =   0
      Top             =   3090
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1323
   End
   Begin VB.Image Image2 
      Height          =   1440
      Left            =   4245
      Picture         =   "frmTPB.frx":101C6
      Top             =   285
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      Top             =   1800
      Width           =   2700
   End
End
Attribute VB_Name = "frmTPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Test form for this approach to a transparent picture box

' Note that this is NOT a true transparent picturebox. Any background graphics that
' are displayed on the picturebox are not clickable. In otherwords, if you place a
' label or image behind the picture box, it will show thru, but will not be interactive

' See the usercontrol for more comments.



Private Sub Check1_Click(Index As Integer)

    If Check1(Index).Value = 0 Then ' not checked
        If Check1(Abs(Index - 1)).Value = 0 Then ' also not checked, so check it
            Check1(Abs(Index - 1)).Value = 1
            Exit Sub
        End If
    Else
        Check1(Abs(Index - 1)).Value = 0    ' uncheck the other checkbox
        SetUpExample Index
    End If
            
End Sub

' some simple test routines
Private Sub Command1_Click()
    ' make form bkcolor change which should update the "transparent" picturebox too
    Me.BackColor = vbBlue
End Sub

Private Sub Command2_Click()
    ' make label behind picbox change, which should update the "transparent" picturebox too
    Label1.AutoSize = True
    Label1.BackStyle = 0
    Label1.Caption = "This label is actually behind the picture box"
End Sub

Private Sub Command3_Click()
    ' remove the form's picture property, which should update the "transparent" picturebox too
    Set Me.Picture = Nothing
End Sub

Private Sub Command4_Click()
    ' move the foxhead image around the back of the picture box,  which should update the "transparent" picturebox too
    Dim X As Long, Y As Long
    X = CLng(Rnd * ScaleX(Picture1.Width, Picture1.ScaleMode, Me.ScaleMode))
    Y = CLng(Rnd * ScaleY(Picture1.Height, Picture1.ScaleMode, Me.ScaleMode))
    
    Image2.Move X, Y
    
End Sub

Private Sub Form_Load()

    Picture1.AutoRedraw = True
    Check1(0).Value = 1
    
End Sub


Private Sub SetUpExample(Index As Integer)

    Dim X As Long, Y As Long

    ' call routine to give us a top,left coordinate to place the usercontrol and
    ' also associate this usercontrol with a control (Picture1).
    ' The routine will calculate the top,left and take into account the control's
    ' borders, if applicable. The usercontrol needs to be placed exactly under the
    ' picturebox's client area (not the area that includes the borders)
    
    
    ' excuse the ScaleX/Y calcs below. This is used only so you can change scalemodes
    ' in this sample project and also maintain the correct coordinates.
    
    If Index = 1 Then ' partial transparency example
    
        Image1.Visible = True
        Set Image1.Picture = Nothing
        
        Picture1.BackColor = vbWhite
        Set Picture1.Picture = Nothing
        Picture1.Cls
        
        ' get the Image1 Top,Left in pixels
        X = ScaleX(Image1.Left, Picture1.ScaleMode, vbPixels)
        Y = ScaleY(Image1.Top, Picture1.ScaleMode, vbPixels)
        ' associate with Image1
        UserControl11.Initialize Me.hwnd, Picture1.hwnd, X, Y, Image1
        ' position usercontrol exactly under Image1
        With Picture1
            UserControl11.Move ScaleX(X, vbPixels, Me.ScaleMode), ScaleY(Y, vbPixels, Me.ScaleMode), _
                ScaleX(Image1.Width, .ScaleMode, Me.ScaleMode), ScaleY(Image1.Height, .ScaleMode, Me.ScaleMode)
        End With
        Set Image1.Picture = UserControl11.Image
        
    Else            ' 100% transparent
    
        Image1.Visible = False
        Picture1.BackColor = vbButtonFace
        ' associate with Picture1
        UserControl11.Initialize Me.hwnd, Picture1.hwnd, X, Y, Picture1
        ' position usercontrol exactly under Picture1, shifting for borders if needed
        With Picture1
            UserControl11.Move ScaleX(X, vbPixels, Me.ScaleMode), ScaleY(Y, vbPixels, Me.ScaleMode), _
                ScaleX(.ScaleWidth, .ScaleMode, Me.ScaleMode), ScaleY(.ScaleHeight, .ScaleMode, Me.ScaleMode)
        End With
        ' set our PictureBox's Picture property to the usercontrol's Image
        Set Picture1.Picture = UserControl11.Image
    End If

    ' ensure our usercontrol is top most; this also triggers a repaint if needed
    UserControl11.ZOrder

End Sub
