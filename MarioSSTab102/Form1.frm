VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SubClassing SSTab "
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option8 
      Caption         =   "AnimatedGradient"
      Height          =   255
      Left            =   7800
      TabIndex        =   34
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   7680
      TabIndex        =   33
      Top             =   6480
      Width           =   3855
      Begin MSComctlLib.Slider Slider1 
         Height          =   555
         Left            =   600
         TabIndex        =   35
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   979
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fade Speed 1-10"
         ForeColor       =   &H00865724&
         Height          =   195
         Left            =   1320
         TabIndex        =   36
         Top             =   840
         Width           =   1230
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   7560
   End
   Begin VB.Frame Frame2 
      Caption         =   "1"
      Height          =   1695
      Left            =   7560
      TabIndex        =   26
      Top             =   3960
      Width           =   1935
      Begin VB.OptionButton Option6 
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H0020A222&
         Height          =   375
         Left            =   1080
         MouseIcon       =   "Form1.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00CDB5A0&
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form1.frx":0152
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   840
         TabIndex        =   31
         Top             =   480
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "2"
      Height          =   1695
      Left            =   9720
      TabIndex        =   20
      Top             =   3960
      Width           =   1935
      Begin VB.OptionButton Option5 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form1.frx":02A4
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00B99D7F&
         Height          =   375
         Left            =   1080
         MouseIcon       =   "Form1.frx":03F6
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   480
         Width           =   195
      End
   End
   Begin VB.OptionButton Option3 
      Caption         =   "GradientColor"
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   3480
      Width           =   2295
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   9960
      MouseIcon       =   "Form1.frx":0548
      MousePointer    =   99  'Custom
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   8520
      MouseIcon       =   "Form1.frx":069A
      MousePointer    =   99  'Custom
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   8520
      MouseIcon       =   "Form1.frx":07EC
      MousePointer    =   99  'Custom
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9960
      MouseIcon       =   "Form1.frx":093E
      MousePointer    =   99  'Custom
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply SubClassing"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   8160
      Width           =   2775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4260
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0A90
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
         Height          =   495
         Left            =   4440
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Alberto Flores Gonzalez"
         Height          =   195
         Left            =   1800
         TabIndex        =   9
         Top             =   1440
         Width           =   2100
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "SolidColor"
      Height          =   195
      Left            =   7680
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Picture Style"
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2655
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0AAC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   4200
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Alberto Flores Gonzalez"
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   2100
      End
   End
   Begin VB.Label Label13 
      Caption         =   $"Form1.frx":0AC8
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   7560
      TabIndex        =   37
      Top             =   7920
      Width           =   4200
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "SSTab Subclassing By Mario Flores G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   32
      Top             =   240
      Width           =   4035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00825623&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B99D7F&
      FillColor       =   &H00825623&
      FillStyle       =   0  'Solid
      Height          =   8895
      Left            =   7320
      Top             =   0
      Width           =   60
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00825623&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   7320
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00825623&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   7320
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00825623&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   7320
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Left            =   10440
      TabIndex        =   19
      Top             =   2280
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Left            =   9000
      TabIndex        =   18
      Top             =   2280
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Left            =   10440
      TabIndex        =   17
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Left            =   9000
      TabIndex        =   16
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ssStyleTabbedDialog"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   7200
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ssStylePropertPage"
      Height          =   195
      Left            =   2520
      TabIndex        =   0
      Top             =   3960
      Width           =   1380
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   -2520
      Picture         =   "Form1.frx":0B79
      Tag             =   "YUST FOR DEMO "
      Top             =   960
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   -2520
      Picture         =   "Form1.frx":292A
      Tag             =   "YUST FOR DEMO "
      Top             =   600
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'SUBCLASSING THE SSTab Control  By Mario Alberto Flores Gonzalez
'version 1.1
'February 10, 2004
'Feel free to use this source code as you wish in your projects

'                        sistec_de_juarez@hotmail.com

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


Option Explicit

 Dim xWidth  As Long
 Dim xHeight As Long
 Dim xColor1 As Long
 Dim xColor2 As Long
 
 Dim gColor1 As Long
 Dim gColor2 As Long
 Dim gColor3 As Long
 Dim gColor4 As Long
 
 Dim gDir    As Direction
 

 
Private Sub Command3_Click()
'=====================================================================
   'SSTab 1
   '=====================================================================
    
    '--------------------------------------------------------------------------------
    'Bitmaped Style
    '--------------------------------------------------------------------------------
   
    If Option1.Value = True Then
      
        SetStyle SSTab1.hWnd, cPicture                  '//--- Set The Style of The SSTab
        ScaleMyPictures Image1                          '//--- Calculate Width & Height of Image
        SetPicture SSTab1.hWnd, xWidth, xHeight, Image1 '//--- Asing new Picture
    
    
    '--------------------------------------------------------------------------------
    'Solid Color Style
    '--------------------------------------------------------------------------------
    
    ElseIf Option2.Value = True Then

        SetStyle SSTab1.hWnd, cSolidColor   '//--- Set The Style of The SSTab
        SetSolidColor SSTab1.hWnd, xColor1  '//--- Asing new Color

    '--------------------------------------------------------------------------------
    'GradientColor Style
    '--------------------------------------------------------------------------------

    ElseIf Option3.Value = True Then
             
        If Option4.Value = True Then
           gDir = cHorizontal
        Else
           gDir = cVertical
        End If
        
        SetStyle SSTab1.hWnd, cGradient  '//--- Set The Style of The SSTab
        SetGradientDir SSTab1.hWnd, gDir '//--- Set The Gradient Direction
        SetGradientColor1 SSTab1.hWnd, gColor1  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab1.hWnd, gColor2  '//--- Asing new Gradient Color End
 
    '--------------------------------------------------------------------------------
    'AnimatedGradient Style
    '--------------------------------------------------------------------------------
        
    ElseIf Option8.Value = True Then
        
             
        If Option4.Value = True Then
           gDir = cHorizontal
        Else
           gDir = cVertical
        End If
        
        SetStyle SSTab1.hWnd, cAnimatedGradient '//--- Set The Style of The SSTab
        SetFadeTime SSTab1.hWnd, Slider1.Value  '//--- Set The Fade Time
        SetGradientDir SSTab1.hWnd, gDir        '//--- Set The Gradient Direction
        SetGradientColor1 SSTab1.hWnd, gColor1  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab1.hWnd, gColor2  '//--- Asing new Gradient Color End
  
    End If
    
        
    SSTabSubclass SSTab1.hWnd '//--- Begin SubClassing
    
   '=====================================================================
   'SSTab 2
   '=====================================================================
    
    
    '--------------------------------------------------------------------------------
    'Bitmaped Style
    '--------------------------------------------------------------------------------
    
    If Option1.Value = True Then
        
        SetStyle SSTab2.hWnd, cPicture                  '//--- Set The Style of The SSTab
        ScaleMyPictures Image2                          '//--- Calculate Width & Height of Image
        SetPicture SSTab2.hWnd, xWidth, xHeight, Image2 '//--- Asing new Picture
     
    '--------------------------------------------------------------------------------
    'SolidColor Style
    '--------------------------------------------------------------------------------
    
    ElseIf Option2.Value = True Then

        SetStyle SSTab2.hWnd, cSolidColor   '//--- Set The Style of The SSTab
        SetSolidColor SSTab2.hWnd, xColor2  '//--- Asing new Color

    '--------------------------------------------------------------------------------
    'GradientColor Style
    '--------------------------------------------------------------------------------

    ElseIf Option3.Value = True Then
        
        If Option6.Value = True Then
           gDir = cHorizontal
        Else
           gDir = cVertical
        End If
        
        SetStyle SSTab2.hWnd, cGradient         '//--- Set The Style of The SSTab
        SetGradientDir SSTab2.hWnd, gDir        '//--- Set The Gradient Direction
        SetGradientColor1 SSTab2.hWnd, gColor3  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab2.hWnd, gColor4  '//--- Asing new Gradient Color End
    
    
    '--------------------------------------------------------------------------------
    'AnimatedGradient Style
    '--------------------------------------------------------------------------------

    ElseIf Option8.Value = True Then
        
        If Option4.Value = True Then
           gDir = cHorizontal
        Else
           gDir = cVertical
        End If
        
        SetStyle SSTab2.hWnd, cAnimatedGradient '//--- Set The Style of The SSTab
        SetFadeTime SSTab2.hWnd, Slider1.Value  '//--- Set The Fade Time
        SetGradientDir SSTab2.hWnd, gDir        '//--- Set The Gradient Direction
        SetGradientColor1 SSTab2.hWnd, gColor3  '//--- Asing new Gradient Color Start
        SetGradientColor2 SSTab2.hWnd, gColor4  '//--- Asing new Gradient Color End
  
    End If

    

     
        
    SSTabSubclass SSTab2.hWnd '//--- Begin SubClassing
        
        
    RedrawWindow SSTab2.hWnd, ByVal 0&, ByVal 0&, &H1
    RedrawWindow SSTab1.hWnd, ByVal 0&, ByVal 0&, &H1
    'Me.Hide
    'Me.Show

End Sub

'=====================================================================
' Little Function To Calculate Width & Height of Image
'=====================================================================

Private Sub ScaleMyPictures(ByRef Pic As StdPicture)
  xWidth = Me.ScaleX(Pic.Width, vbHimetric, vbPixels)
  xHeight = Me.ScaleY(Pic.Height, vbHimetric, vbPixels)
End Sub


Private Sub Form_Load()
Picture1.Picture = Image1.Picture
Picture2.Picture = Image2.Picture

xColor1 = Picture4.BackColor
xColor2 = Picture3.BackColor

    gColor1 = Picture5.BackColor
    gColor2 = Picture6.BackColor
    gColor3 = Picture7.BackColor
    gColor4 = Picture8.BackColor

End Sub

Private Sub Picture1_Click()
CD.DialogTitle = "Select a Picture File"
CD.Filter = "JPG|*.JPG|BMP|*.BMP"
CD.ShowOpen

    If Len(Trim(CD.FileName)) > 0 Then
        Picture1.Picture = LoadPicture(CD.FileName)
        Image1.Picture = Picture1.Picture
    End If


End Sub

Private Sub Picture2_Click()
CD.DialogTitle = "Select a Picture File"
CD.Filter = "JPG|*.JPG|BMP|*.BMP"
CD.ShowOpen

    If Len(Trim(CD.FileName)) > 0 Then
        Picture2.Picture = LoadPicture(CD.FileName)
        Image2.Picture = Picture2.Picture
    End If

End Sub

Private Sub Picture3_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture3.BackColor = CD.Color
        xColor2 = CD.Color
    End If

End Sub

Private Sub Picture4_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture4.BackColor = CD.Color
        xColor1 = CD.Color
    End If

End Sub

Private Sub Picture5_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture5.BackColor = CD.Color
        gColor1 = CD.Color
    End If

End Sub

Private Sub Picture6_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture6.BackColor = CD.Color
        gColor2 = CD.Color
    End If

End Sub

Private Sub Picture7_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture7.BackColor = CD.Color
        gColor3 = CD.Color
    End If

End Sub

Private Sub Picture8_Click()
CD.ShowColor

    If Val(CD.Color) > 0 Then
        Picture8.BackColor = CD.Color
        gColor4 = CD.Color
    End If

End Sub




