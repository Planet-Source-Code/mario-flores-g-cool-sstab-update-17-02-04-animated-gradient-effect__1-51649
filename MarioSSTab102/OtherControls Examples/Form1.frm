VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2655
      Left            =   2880
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   4
      Top             =   4560
      Width           =   3375
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   1320
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1191
      _Version        =   393216
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19660801
      CurrentDate     =   38028
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19660801
      CurrentDate     =   38028
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   675
      Left            =   6120
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1191
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   240
      Picture         =   "Form1.frx":24B8
      Top             =   4080
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        SetStyle MonthView1.hwnd, cPicture
        ScaleMyPictures Image1
        SetPicture MonthView1.hwnd, xWidth, xHeight, Image1
        
        SSTabSubclass MonthView1.hwnd
         
        SetStyle Slider1.hwnd, cPicture
        ScaleMyPictures Image1
        SetPicture Slider1.hwnd, xWidth, xHeight, Image1

        SSTabSubclass Slider1.hwnd
       
        SetStyle MonthView2.hwnd, cGradient
        SetGradientDir MonthView2.hwnd, cVertical
        SetGradientColor1 MonthView2.hwnd, vbYellow
        SetGradientColor2 MonthView2.hwnd, vbGreen
        
        SSTabSubclass MonthView2.hwnd
         
        SetStyle Slider2.hwnd, cGradient
        SetGradientDir Slider2.hwnd, cHorizontal
        SetGradientColor1 Slider2.hwnd, vbRed
        SetGradientColor2 Slider2.hwnd, vbYellow
    
        SSTabSubclass Slider2.hwnd
      
        SetStyle MSChart1.hwnd, cPicture
        ScaleMyPictures Image1
        SetPicture MSChart1.hwnd, xWidth, xHeight, Image1

        SSTabSubclass MSChart1.hwnd
     
        

End Sub
'=====================================================================
' Little Function To Calculate Width & Height of Image
'=====================================================================

Private Sub ScaleMyPictures(ByRef Pic As StdPicture)
  xWidth = Me.ScaleX(Pic.Width, vbHimetric, vbPixels)
  xHeight = Me.ScaleY(Pic.Height, vbHimetric, vbPixels)
End Sub

