VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{305A9443-6DC0-101D-85F5-6EBA1EE93AF4}#1.1#0"; "HINDIC32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Mp3 Player 2"
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7470
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5760
      Picture         =   "Form1.frx":0CCA
      ScaleHeight     =   180
      ScaleWidth      =   1695
      TabIndex        =   43
      Top             =   1680
      Width           =   1700
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "Form1.frx":1D5C
      ScaleHeight     =   180
      ScaleWidth      =   1695
      TabIndex        =   42
      Top             =   1680
      Width           =   1700
   End
   Begin VB.CommandButton cmdPlaylistT 
      Caption         =   "Playlist ?"
      Height          =   255
      Left            =   1560
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrPause2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   6480
   End
   Begin VB.Timer tmrPause1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   6480
   End
   Begin VB.PictureBox pctMin 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   7020
      Picture         =   "Form1.frx":2DEE
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   38
      Top             =   75
      Width           =   135
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   480
      Width           =   2775
   End
   Begin RichTextLib.RichTextBox rtbId 
      Height          =   495
      Left            =   6120
      TabIndex        =   36
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":2F2C
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   2520
      Picture         =   "Form1.frx":2FDA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctStopD 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4080
      Picture         =   "Form1.frx":331C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctPauseD 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   3720
      Picture         =   "Form1.frx":365E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctPlayD 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   3360
      Picture         =   "Form1.frx":39A0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctFFD 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4440
      Picture         =   "Form1.frx":3CE2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctRRD 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   3000
      Picture         =   "Form1.frx":4024
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   1150
      Width           =   235
   End
   Begin VB.PictureBox pctLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H0021BA9C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      Picture         =   "Form1.frx":4366
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   1150
      Width           =   235
   End
   Begin VB.CommandButton cmdNextList 
      Caption         =   "Advance List"
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   26
      Text            =   "Text3"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrVU 
      Interval        =   50
      Left            =   960
      Top             =   6000
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Use Playlist"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlaylist 
      Caption         =   "Playlist"
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   145
      Width           =   2760
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2520
      MultiSelect     =   2  'Extended
      TabIndex        =   17
      Top             =   8520
      Width           =   1815
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2085
      ItemData        =   "Form1.frx":46D8
      Left            =   5160
      List            =   "Form1.frx":46DA
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   5280
      Width           =   3510
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2400
      Top             =   6000
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1920
      Tag             =   "0"
      Top             =   6000
   End
   Begin ComctlLib.Slider HScroll2 
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "Form1.frx":46DC
      LargeChange     =   1000
      SmallChange     =   50
      Min             =   -5000
      Max             =   5000
      TickStyle       =   3
   End
   Begin ComctlLib.Slider HScroll1 
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "Form1.frx":49F6
      LargeChange     =   500
      SmallChange     =   10
      Max             =   2500
      TickStyle       =   3
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Quit"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   6000
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Pause"
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Play"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   6975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   0
      Picture         =   "Form1.frx":4D10
      ScaleHeight     =   1215
      ScaleWidth      =   3135
      TabIndex        =   22
      Top             =   0
      Width           =   3135
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      MouseIcon       =   "Form1.frx":F612
      LargeChange     =   1
      Max             =   1
      TickStyle       =   3
   End
   Begin HindicLib.HIndicator HIndicator1 
      Height          =   135
      Left            =   3360
      TabIndex        =   40
      Top             =   360
      Width           =   1845
      _Version        =   65537
      _ExtentX        =   3254
      _ExtentY        =   238
      _StockProps     =   65
      BackColor       =   0
      Border          =   1
      BevelOuter      =   1
      BevelWidth      =   1
      ItemBackColor   =   0
      ItemCount1      =   7
      ItemCount2      =   3
      ItemCount3      =   2
      ItemForeColor1  =   65280
      ItemForeColor2  =   65535
      ItemForeColor3  =   255
      BorderWidth     =   1
      BevelInner      =   0
      Value           =   100
      Max             =   100
      Min             =   0
   End
   Begin HindicLib.HIndicator HIndicator2 
      Height          =   135
      Left            =   3360
      TabIndex        =   41
      Top             =   600
      Width           =   1845
      _Version        =   65537
      _ExtentX        =   3254
      _ExtentY        =   238
      _StockProps     =   65
      BackColor       =   0
      Border          =   1
      BevelOuter      =   1
      BevelWidth      =   1
      ItemBackColor   =   0
      ItemCount1      =   7
      ItemCount2      =   3
      ItemCount3      =   2
      ItemForeColor1  =   65280
      ItemForeColor2  =   65535
      ItemForeColor3  =   255
      BorderWidth     =   1
      BevelInner      =   0
      Value           =   100
      Max             =   100
      Min             =   0
   End
   Begin VB.Label lblMute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mute"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   840
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   7200
      Picture         =   "Form1.frx":F92C
      Top             =   75
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      Height          =   135
      Left            =   -360
      TabIndex        =   23
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3600
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   1150
      Width           =   945
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   4080
      TabIndex        =   15
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5805
      TabIndex        =   8
      ToolTipText     =   "Double click me to set volume to 100%."
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      ToolTipText     =   "Double click me to set volume to 50%."
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "Double click me to return balance to center !"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5805
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   1095
      Left            =   -1500
      TabIndex        =   2
      Top             =   360
      Width           =   3135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   4210752
      DisplayForeColor=   255
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private MP3 As clsMP3
Dim Mini As Boolean
Dim Pau As Boolean
Dim Title As String
Dim Onestep As Boolean
Dim theta(20) As Single
Dim rates(20) As Single
Dim hmixer As Long                  ' mixer handle
Dim inputVolCtrl As MIXERCONTROL    ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL   ' microphone volume control
Dim rc As Long                      ' return code
Dim ok As Boolean                   ' boolean return code

Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' handle to volume memory
   Private Function fMakeATranspArea(AreaType As String, pCordinate() As Long) As Boolean


       'Name: fMakeATranpArea
       'Author: Dalin Nie
       'Date: 5/18/98
       'Purpose: Create a Transprarent Area in a form so that you can se
       '     e through
       'Input: Areatype : a String indicate what kind of hole shape it w
       '     ould like to make
       ' PCordinate : the cordinate area needed for create the shape:
       ' Example: X1, Y1, X2, Y2 for Rectangle
       'OutPut: A boolean
       Const RGN_DIFF = 4
       Dim lOriginalForm As Long
       Dim ltheHole As Long
       Dim lNewForm As Long
       Dim lFwidth As Single
       Dim lFHeight As Single
       Dim lborder_width As Single
       Dim ltitle_height As Single
       On Error GoTo Trap
       lFwidth = ScaleX(Width, vbTwips, vbPixels)
       lFHeight = ScaleY(Height, vbTwips, vbPixels)
       lOriginalForm = CreateRectRgn(0, 0, lFwidth, lFHeight)
       
       lborder_width = (lFHeight - ScaleWidth) / 2
       ltitle_height = lFHeight - lborder_width - ScaleHeight


       Select Case AreaType
           
           Case "Elliptic"
           
           ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
           Case "RectAngle"
           
           ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
           
           Case "RoundRect"
           
           ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
           Case "Circle"
           ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
           
           Case Else
           MsgBox "Unknown Shape!!"
           Exit Function
       End Select


   lNewForm = CreateRectRgn(0, 0, 0, 0)
   CombineRgn lNewForm, lOriginalForm, _
   ltheHole, RGN_DIFF

   SetWindowRgn hwnd, lNewForm, True
   Me.Refresh
   fMakeATranspArea = True
   Exit Function
Trap:
   MsgBox "error Occurred. Error # " & err.Number & ", " & err.Description



   End Function




Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.Text = frmPlay.txtOut.Text
End If
End Sub

Private Sub cmdNextList_Click()
    MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & _
        "Artist: " & MP3.Artist & vbCrLf & _
        "Album: " & MP3.Album & vbCrLf
        txtId.Text = rtbId.Text

    Title = String(30, " ") + "Mp3 Player 2" & " (Playing..." & MP3.Artist & ")"
    frmMain.Caption = "Mp3 Player 2" & " " & MP3.Title & " " & "-" & " " & MP3.Artist
Slider1.Max = MediaPlayer1.Duration

End Sub

Private Sub cmdPlaylistT_Click()
MediaPlayer1.Filename = Text1.Text
    MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & "Artist: " & MP3.Artist & vbCrLf & "Album: " & MP3.Album & vbCrLf
rtbId.Text = "Track: " & MP3.Title & vbCrLf & "Artist: " & MP3.Artist & vbCrLf & "Album: " & MP3.Album & vbCrLf
txtId.Text = rtbId.Text
Title = String(30, " ") + "Mp3 Player 2" & " (Playlist loaded..." & MP3.Artist & " on deck..." & ")"
frmMain.Caption = "Mp3 Player 2" & " " & "(" & MP3.Title & " " & "-" & " " & MP3.Artist & ")"
End Sub

Private Sub Command1_Click()
Command2.Visible = True
Command5.Visible = False
rtbId.Refresh
txtId.Text = rtbId.Text


       On Error GoTo ErrHandler
       CommonDialog1.CancelError = True
       CommonDialog1.DialogTitle = "Select a Mp3."
       CommonDialog1.Filter = "MP3 Files(*.mp3)|*.mp3"
       CommonDialog1.ShowOpen


       If CommonDialog1.Filename <> "" Then
           'MediaPlayer1.FileName = CommonDialog1.FileName
           MediaPlayer1.Filename = Text1.Text
           MediaPlayer1.AutoStart = False
           MediaPlayer1.Stop
           Text1.Text = CommonDialog1.Filename
           End If
                  MP3.Filename = RTrim(Text1.Text)
                  rtbId.Text = "Track: " & MP3.Title & vbCrLf & "Artist: " & MP3.Artist & vbCrLf & "Album: " & MP3.Album & vbCrLf
                  txtId.Text = rtbId.Text
                  Pau = True
                  MediaPlayer1.Filename = Text1.Text
                  MediaPlayer1.Play
                  Slider1.Max = MediaPlayer1.Duration
                  Title = String(30, " ") + "Mp3 Player 2" & " (Playing..." & MP3.Artist & ")"
                  frmMain.Caption = "Mp3 Player 2" & " " & "(" & MP3.Title & " " & "-" & " " & MP3.Artist & ")"
                  frmPlay.pctPlay.Picture = frmPlay.pctOn.Picture
                  frmPlay.pctCancel.Picture = frmPlay.pctOff.Picture
                  frmPlay.txtOut.Text = Text1.Text

           Exit Sub
ErrHandler:
           'If an error occurs, exit the procedure.
End Sub

Private Sub Command10_Click()
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
    Title = String(30, " ") + "Mp3 Player 2" & " (No Artist...)"
    frmPlay.pctCancel.Picture = frmPlay.pctOn.Picture
    frmPlay.pctPlay.Picture = frmPlay.pctOff.Picture
    rtbId.Text = "Please select a MP3 file."
    txtId.Text = rtbId.Text

Exit Sub
End If
    MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & _
        "Artist: " & MP3.Artist & vbCrLf & _
        "Album: " & MP3.Album & vbCrLf
txtId.Text = rtbId.Text

Pau = True

    MediaPlayer1.Filename = Text1.Text
    MediaPlayer1.Play

    Slider1.Max = MediaPlayer1.Duration
    Title = String(30, " ") + "Mp3 Player 2" & " (Playing..." & MP3.Artist & ")"
    frmMain.Caption = "Mp3 Player 2" & " " & "(" & MP3.Title & " " & "-" & " " & MP3.Artist & ")"
    Command2.Visible = False
    Command5.Visible = True
End Sub

Private Sub Command4_Click()
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Command5_Click()
   MediaPlayer1.Stop
   frmPlay.pctCancel.Picture = frmPlay.pctOn.Picture
   frmPlay.pctPlay.Picture = frmPlay.pctOff.Picture
   Title = String(30, " ") + "Mp3 Player 2" & " (Stopped...)"
   Command2.Visible = True
   Command5.Visible = False

End Sub

Private Sub Command6_Click()
    MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & _
        "Artist: " & MP3.Artist & vbCrLf & _
        "Album: " & MP3.Album & vbCrLf
        txtId.Text = rtbId.Text
tmrPause1.Enabled = False
tmrPause2.Enabled = False
frmPlay.pctPlay.Picture = frmPlay.pctOn.Picture
   MediaPlayer1.Play

    Title = String(30, " ") + "Mp3 Player 2" & " (Playing..." & MP3.Artist & ")"
    frmMain.Caption = "Mp3 Player 2" & " " & MP3.Title & " " & "-" & " " & MP3.Artist
Slider1.Max = MediaPlayer1.Duration
   Command6.Visible = False
   Command5.Visible = True
End Sub

Private Sub Command7_Click()
If Pau Then
   MediaPlayer1.Pause
   tmrPause1.Enabled = True
   Title = String(30, " ") + "Mp3 Player 2" & " (Paused...)"
   Command5.Visible = False
   Command6.Visible = True
End If
End Sub

Private Sub Command8_Click()
Unload Me
End
End Sub

Private Sub Command9_Click()
End Sub

Private Sub Form_Load()
Dim lParam(1 To 6) As Long
Dim nRet As Long
Dim i As Integer
Dim Playlst As Variant, intPlaylst As Integer
Me.Left = 4000
Me.Top = 3000
Set MP3 = New clsMP3
Onestep = False
Pau = False
Mini = False
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
       
   ' Get the output volume meter
   ok = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   
   If (ok = True) Then
      HIndicator1.Min = 0
      HIndicator1.Max = outputVolCtrl.lMaximum
      HIndicator2.Max = outputVolCtrl.lMaximum
      
   Else
      HIndicator1.Enabled = False
      HIndicator1.Value = HIndicator1.Min
      HIndicator2.Enabled = False
      HIndicator2.Value = HIndicator1.Min
      MsgBox "Could not initialize level indicators.", vbInformation, "Mp3 Player 2"
   End If
   
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   lParam(1) = 100   'position from left
   lParam(2) = 110   'height from top
   lParam(3) = 400   'width left to right
   lParam(4) = 200  'distance from top?
   lParam(5) = 50 'Undefined
   lParam(6) = 50 'Undefined
   'Call fMakeATranspArea("RoundRect", lParam())
   Call fMakeATranspArea("RectAngle", lParam())
   'Call fMakeATranspArea("Circle", lParam())
   'Call fMakeATranspArea("Elliptic", lParam())

Load frmPlay
frmPlay.Left = frmMain.Left
frmPlay.Top = frmMain.Top + frmMain.Height

Playlst = GetAllSettings(appName:="Mp3Player2", Section:="Playlist")
If Not IsEmpty(Playlst) Then
Text3.Text = Playlst(intPlaylst, 1)
End If
If Text3.Text = "True" Then
frmPlay.Show
ElseIf Text3.Text = "False" Then
frmPlay.Hide
End If
Timer2.Enabled = False
    Randomize
    For i = 0 To 19
        theta(i) = Rnd(1) * 0.14159
        rates(i) = 0.5 * Rnd(1)
    Next

   MediaPlayer1.ShowAudioControls = True
   MediaPlayer1.ShowControls = True
   MediaPlayer1.ShowPositionControls = True
   Command2.Visible = True
   Command5.Visible = False
   Command6.Visible = False
   Label2.Caption = Label2.Caption & " %"
   HScroll1.Value = 1875
   Label2.ForeColor = &HFF&
   If Text1.Text > "" Then
    Title = String(30, " ") + "Mp3 Player 2" & " (Using playlist...Press play...)"
       MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & _
        "Artist: " & MP3.Artist & vbCrLf & _
        "Album: " & MP3.Album & vbCrLf
        txtId.Text = rtbId.Text

   Check1.Value = 1
   Timer1.Enabled = True
   ElseIf Text1.Text = "" Then
   Title = String(30, " ") + "Mp3 Player 2 (Waiting...)"
   End If
If HIndicator1.Value = 0 Then
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    frmPlay.Left = Me.Left
    frmPlay.Top = Me.Top + Me.Height
End If

Label2.ForeColor = &HFF&
Label4.ForeColor = &HFF&

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmPlay
End Sub

Private Sub Form_Resize()
Dim X As Variant
Dim Playlst As Variant, intPlaylst As Integer
Playlst = GetAllSettings(appName:="Mp3Player2", Section:="Playlist")
If Not IsEmpty(Playlst) Then
Text3.Text = Playlst(intPlaylst, 1)
End If
If Text3.Text = "True" Then
frmPlay.WindowState = vbNormal
End If
If Text3.Text = "True" Then
If Me.WindowState = 1 Then
frmPlay.Hide
ElseIf Me.WindowState = 0 Then
X = ShowWindow(frmPlay.hwnd, 1) ' Where frmID is the name Property of the MDI child form you want to hide.
End If
End If

End Sub

Private Sub HScroll1_Change()
Dim pim, sha
Dim foo As Integer, poo As Integer
Label2.ForeColor = &HFFFF&
sha = HScroll1.Value - 2500
MediaPlayer1.volume = sha
On Error GoTo hell
poo = HScroll1.Min
foo = HScroll1.Value
Label2.Caption = foo \ 25 & " %"
hell:
Exit Sub

End Sub

Private Sub HScroll1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF&
End Sub

Private Sub HScroll1_Scroll()
Dim pim, sha
Dim foo As Integer, poo As Integer
Label2.ForeColor = &HFFFF&
sha = HScroll1.Value - 2500
MediaPlayer1.volume = sha
On Error GoTo hell
poo = HScroll1.Min
foo = HScroll1.Value
Label2.Caption = foo \ 25 & " %"
hell:
Exit Sub

End Sub

Private Sub HScroll2_Change()
On Error GoTo hell
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Label4.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Label4.Caption = "Left"
End If
If HScroll2.Value > 500 Then
Label4.Caption = "Right"
End If
MediaPlayer1.Balance = HScroll2.Value
hell:
Exit Sub


End Sub

Private Sub HScroll2_Scroll()
On Error GoTo hell
Label4.ForeColor = &HFF&
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Label4.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Label4.Caption = "Left"
End If
If HScroll2.Value > 500 Then
Label4.Caption = "Right"
End If
MediaPlayer1.Balance = HScroll2.Value
Label4.ForeColor = &HFFFF&
hell:
Exit Sub

End Sub

Private Sub Image2_Click()
Unload Me
End
End Sub

Private Sub Image2_DblClick()
Image2.Left = Image2.Left + 15
Image2.Top = Image2.Top + 15
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Left = Image2.Left + 15
Image2.Top = Image2.Top + 15
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Left = Image2.Left - 15
Image2.Top = Image2.Top - 15
End Sub

Private Sub Label1_DblClick()
HScroll1.Value = 2500
End Sub

Private Sub Label2_DblClick()
HScroll1.Value = 1250
End Sub

Private Sub Label4_DblClick()
HScroll2.Value = 0

End Sub

Private Sub lblClose_Click()
End Sub


Private Sub lblMute_DblClick()
lblMute.Left = lblMute.Left + 15
lblMute.Top = lblMute.Top + 15
End Sub

Private Sub lblMute_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMute.Left = lblMute.Left + 15
lblMute.Top = lblMute.Top + 15
End Sub

Private Sub lblMute_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMute.Left = lblMute.Left - 15
lblMute.Top = lblMute.Top - 15
If MediaPlayer1.Mute = False Then
MediaPlayer1.Mute = True
lblMute.ForeColor = &HFFFF&
   HScroll1.Enabled = False
   Label2.Caption = "Mute"

ElseIf MediaPlayer1.Mute = True Then
MediaPlayer1.Mute = False
lblMute.ForeColor = &HFF&
   HScroll1.Enabled = True
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.Min
foo = HScroll1.Value
Label2.Caption = foo \ 25 & " %"
hell:
Exit Sub
End If
End Sub

Private Sub pctFFD_DblClick()
pctFFD.Left = pctFFD.Left + 15
pctFFD.Top = pctFFD.Top + 15
End Sub

Private Sub pctFFD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctFFD.Left = pctFFD.Left + 15
pctFFD.Top = pctFFD.Top + 15
End Sub

Private Sub pctFFD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctFFD.Left = pctFFD.Left - 15
pctFFD.Top = pctFFD.Top - 15
If Check1.Value = 0 Then
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 10
ElseIf Check1.Value = 1 Then
frmPlay.cmdNext = True
End If
End Sub

Private Sub pctLoad_DblClick()
pctLoad.Left = pctLoad.Left + 15
pctLoad.Top = pctLoad.Top + 15
End Sub

Private Sub pctLoad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctLoad.Left = pctLoad.Left + 15
pctLoad.Top = pctLoad.Top + 15
End Sub

Private Sub pctLoad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctLoad.Left = pctLoad.Left - 15
pctLoad.Top = pctLoad.Top - 15
Command1 = True
End Sub

Private Sub pctMin_DblClick()
pctMin.Left = pctMin.Left + 15
pctMin.Top = pctMin.Top + 15
End Sub

Private Sub pctMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctMin.Left = pctMin.Left + 15
pctMin.Top = pctMin.Top + 15
End Sub

Private Sub pctMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctMin.Left = pctMin.Left - 15
pctMin.Top = pctMin.Top - 15
Mini = True
Me.WindowState = vbMinimized
End Sub

Private Sub pctPauseD_DblClick()
pctPauseD.Left = pctPauseD.Left + 15
pctPauseD.Top = pctPauseD.Top + 16
End Sub

Private Sub pctPauseD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPauseD.Left = pctPauseD.Left + 15
pctPauseD.Top = pctPauseD.Top + 15
End Sub

Private Sub pctPauseD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPauseD.Left = pctPauseD.Left - 15
pctPauseD.Top = pctPauseD.Top - 15
If Command6.Visible = True Then
Command6 = True
ElseIf Command6.Visible = False Then
Command7 = True
End If
End Sub

Private Sub pctPlayD_DblClick()
pctPlayD.Left = pctPlayD.Left + 15
pctPlayD.Top = pctPlayD.Top + 15
End Sub

Private Sub pctPlayD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPlayD.Left = pctPlayD.Left + 15
pctPlayD.Top = pctPlayD.Top + 15
End Sub

Private Sub pctPlayD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPlayD.Left = pctPlayD.Left - 15
pctPlayD.Top = pctPlayD.Top - 15
frmPlay.pctPlay.Picture = frmPlay.pctOn.Picture
frmPlay.pctCancel.Picture = frmPlay.pctOff.Picture
If Text1.Text = "" Then
    Title = String(30, " ") + "Mp3 Player 2" & " (No Artist...)"
    rtbId.Text = "Please select a MP3 file."
    txtId.Text = rtbId.Text
    frmPlay.pctPlay.Picture = frmPlay.pctOff.Picture
    frmPlay.pctCancel.Picture = frmPlay.pctOn.Picture
Exit Sub
End If

    MP3.Filename = RTrim(Text1.Text)
    rtbId.Text = "Track: " & MP3.Title & vbCrLf & _
        "Artist: " & MP3.Artist & vbCrLf & _
        "Album: " & MP3.Album & vbCrLf
txtId.Text = rtbId.Text

Pau = True

    MediaPlayer1.Filename = Text1.Text
    MediaPlayer1.Play
    tmrPause1.Enabled = False
    tmrPause2.Enabled = False
    Slider1.Max = MediaPlayer1.Duration
    Title = String(30, " ") + "Mp3 Player 2" & " (Playing..." & MP3.Artist & ")"
    frmMain.Caption = "Mp3 Player 2" & " " & "(" & MP3.Title & " " & "-" & " " & MP3.Artist & ")"
    Command2.Visible = False
    Command5.Visible = True
End Sub

Private Sub pctRRD_DblClick()
pctRRD.Left = pctRRD.Left + 15
pctRRD.Top = pctRRD.Top + 15
End Sub

Private Sub pctRRD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctRRD.Left = pctRRD.Left + 15
pctRRD.Top = pctRRD.Top + 15
End Sub

Private Sub pctRRD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctRRD.Left = pctRRD.Left - 15
pctRRD.Top = pctRRD.Top - 15
If MediaPlayer1.CurrentPosition > 1 Then
If Check1.Value = 0 Then
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 10
ElseIf Check1.Value = 1 Then
frmPlay.cmdTrigBack = True
End If
End If
End Sub

Private Sub pctStopD_DblClick()
pctStopD.Left = pctStopD.Left + 15
pctStopD.Top = pctStopD.Top + 15
End Sub

Private Sub pctStopD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctStopD.Left = pctStopD.Left + 15
pctStopD.Top = pctStopD.Top + 15
End Sub

Private Sub pctStopD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctStopD.Left = pctStopD.Left - 15
pctStopD.Top = pctStopD.Top - 15
Command5 = True
End Sub

Private Sub Picture5_DblClick()
Picture5.Left = Picture5.Left + 15
Picture5.Top = Picture5.Top + 15

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Left = Picture5.Left + 15
Picture5.Top = Picture5.Top + 15

End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Left = Picture5.Left - 15
Picture5.Top = Picture5.Top - 15
If frmPlay.Visible = False Then
frmPlay.Left = frmMain.Left
frmPlay.Top = frmMain.Top + frmMain.Height
frmPlay.Show
SaveSetting appName:="Mp3Player2", Section:="Playlist", Key:="Visible", setting:=True
ElseIf frmPlay.Visible = True Then
frmPlay.Hide
SaveSetting appName:="Mp3Player2", Section:="Playlist", Key:="Visible", setting:=False
End If
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
End Sub

Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value

End Sub

Private Sub Timer5_Timer()

End Sub

Private Sub Timer1_Timer()
Slider1.Value = MediaPlayer1.CurrentPosition
If frmPlay.txtCount.Text = frmPlay.txtCurSong.Text Then
Exit Sub
End If
If MediaPlayer1.CurrentPosition = MediaPlayer1.Duration Then
If Check1.Value = 1 Then
If Not Onestep Then
Onestep = True
frmPlay.chkStep.Value = 0
frmPlay.cmdTrig = True
Text1.Text = frmPlay.txtOut.Text
Command2 = True
End If
End If
End If
If MediaPlayer1.CurrentPosition = 0 Then
frmPlay.chkStep.Value = 0
Onestep = False
End If
End Sub

Private Sub Timer2_Timer()
Dim z As Integer
On Error Resume Next
        Title = Mid(Title, 2) & Left(Title, 1)
     
        For z = 0 To 5
        Next z

End Sub

Private Sub Timer3_Timer()
If Timer2.Enabled = False Then
Title = Mid(Title, 2) & Left(Title, 1)
txtTitle.Text = Title
End If

    


End Sub

Private Sub tmrIcon1_Timer()

End Sub

Private Sub Timer4_Timer()

End Sub

Private Sub tmrPause1_Timer()
frmPlay.pctPlay.Picture = frmPlay.pctOff.Picture
tmrPause2.Enabled = True
tmrPause1.Enabled = False
End Sub

Private Sub tmrPause2_Timer()
frmPlay.pctPlay.Picture = frmPlay.pctOn.Picture
tmrPause1.Enabled = True
tmrPause2.Enabled = False
End Sub

Private Sub tmrVU_Timer()

   On Error Resume Next

   ' Process sound buffer if recording
   If (fRecording) Then
      For i = 0 To (NUM_BUFFERS - 1)
         If inHdr(i).dwFlags And WHDR_DONE Then
            rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
         End If
      Next
   End If

   ' Get the current output level
   If (HIndicator1.Enabled = True) Then
      mxcd.dwControlID = outputVolCtrl.dwControlID
      mxcd.item = outputVolCtrl.cMultipleItems
      rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
      
      If (volume < 0) Then volume = -volume
   
      HIndicator1.Value = volume
      HIndicator2.Value = HIndicator1.Value
   End If
   End Sub
