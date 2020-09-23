VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlay 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Mp3 Player 2"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "Playlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctOn 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6960
      Picture         =   "Playlist.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   235
   End
   Begin VB.PictureBox pctOff 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6600
      Picture         =   "Playlist.frx":01FE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   235
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6120
      Picture         =   "Playlist.frx":03F0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   1920
      Width           =   235
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6120
      Picture         =   "Playlist.frx":05E2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   235
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6120
      Picture         =   "Playlist.frx":07D4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   31
      Top             =   1440
      Width           =   235
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6120
      Picture         =   "Playlist.frx":09C6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      Top             =   1200
      Width           =   235
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   6120
      Picture         =   "Playlist.frx":0BB8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   960
      Width           =   235
   End
   Begin VB.PictureBox pctCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":0DAA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   2400
      Width           =   235
   End
   Begin VB.PictureBox pctPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":0F9C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   2160
      Width           =   235
   End
   Begin VB.PictureBox pctBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":118E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   1920
      Width           =   235
   End
   Begin VB.PictureBox pctNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":1380
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   235
   End
   Begin VB.PictureBox pctClearAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":1572
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   235
   End
   Begin VB.PictureBox pctClear1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":1764
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   1200
      Width           =   235
   End
   Begin VB.PictureBox pctEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   235
      Left            =   4560
      Picture         =   "Playlist.frx":1956
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   960
      Width           =   235
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   255
      Left            =   1605
      TabIndex        =   21
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdTrigBack 
      Caption         =   "Back"
      Height          =   255
      Left            =   1665
      TabIndex        =   20
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkStep 
      Caption         =   "Flag"
      Height          =   495
      Left            =   1425
      TabIndex        =   12
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   1425
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrig 
      Caption         =   "Trig"
      Height          =   255
      Left            =   1425
      TabIndex        =   10
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtCount 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   390
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2760
      ItemData        =   "Playlist.frx":1B48
      Left            =   0
      List            =   "Playlist.frx":1B4A
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1815
      Top             =   8235
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Playlist.frx":1B4C
      Top             =   7320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   -195
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8370
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   5000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Total Count"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1425
      TabIndex        =   6
      Top             =   8370
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Timer Count"
      Height          =   735
      Left            =   1305
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCurSong 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   390
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Current Song"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblPlay 
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblClearAll 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear All"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblClear1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Selected"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblNext 
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Playlist"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Title As String
Dim Onestep As Boolean
Sub LoadList(Lst As ListBox, file As String)
Dim A As String
Dim X As String
On Error GoTo Error
Open file For Input As #1
Do Until EOF(1)
Input #1, A$
Lst.AddItem A$
Loop
Close 1
Exit Sub
Error:
X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub cmdCancel_Click()
End Sub

Private Sub cmdClearAll_Click()
End Sub

Private Sub cmdClear_Click()

End Sub

Private Sub cmdDone_Click()
End Sub

Private Sub cmdNext_Click()
If List1.ListIndex < List1.ListCount - 1 Then
List1.ListIndex = List1.ListIndex + 1
txtCurSong.Text = List1.ListIndex + 1
frmMain.MediaPlayer1.Stop
frmMain.Text1.Text = List1.Text
frmMain.MediaPlayer1.Filename = List1.Text
frmMain.Command6 = True
ElseIf List1.ListIndex = List1.ListCount Then
txtCurSong.Text = List1.ListCount
End If
End Sub

Private Sub cmdTrig_Click()
If chkStep.Value = 0 Then
If txtCount.Text = "1" Then
Exit Sub
End If
If Timer1.Interval = 1 Then
If Not txtTime = txtCount - 1 Then
txtTime.Text = txtTime.Text + 1
List1.ListIndex = txtTime.Text
txtOut.Text = List1.Text
txtCurSong.Text = txtCurSong.Text + 1
Else
Timer1.Enabled = False
End If
End If
If txtCount = txtCurSong Then Check1.Value = 1
End If
chkStep.Value = 1
End Sub

Private Sub Command1_Click()

End Sub


Function StripItem(startStrg As String, parser As String) As String
'this takes a string separated by the chr passed in Parser,
'splits off 1 item, and shortens startStrg so that the next
'item is ready for removal.

   Dim c As Integer
   Dim item As String
   
   c = 1
   
   Do
   
      If Mid(startStrg, c, 1) = parser Then
      
         item = Mid(startStrg, 1, c - 1)
         startStrg = Mid(startStrg, c + 1, Len(startStrg))
         StripItem = item
         Exit Function
      End If
      
      c = c + 1
   
   Loop

End Function

'--end block--'
   
Private Sub Command2_Click()
End Sub

Private Sub cmdTrigBack_Click()
If List1.ListIndex > 0 Then
List1.ListIndex = List1.ListIndex - 1
txtCurSong.Text = List1.ListIndex + 1
frmMain.MediaPlayer1.Stop
frmMain.Text1.Text = List1.Text
frmMain.MediaPlayer1.Filename = List1.Text
frmMain.Command6 = True
End If
End Sub

Private Sub Form_Load()
chkStep.Value = 0
pctCancel.Picture = pctOn.Picture
Picture1.Picture = pctOff.Picture
Picture8.Picture = pctOff.Picture
Picture9.Picture = pctOff.Picture
Picture10.Picture = pctOff.Picture
Picture11.Picture = pctOff.Picture
Timer1.Enabled = False
frmPlay.Left = frmMain.Left
frmPlay.Top = frmMain.Top + frmMain.Height
Call LoadList(List1, "C:\Mp3Player2.ini")
       txtCount.Text = List1.ListCount

If txtCount.Text = "0" Then
List1.AddItem "No entries..."
txtCurSong.Text = "0"
txtOut.Text = List1.Text
Check1.Value = 0
Exit Sub
End If
       If txtCount.Text > "0" Then
       txtTime.Text = "0"
       Timer1.Enabled = False
       List1.ListIndex = "0"
       frmMain.Text1.Text = List1.Text
       txtOut.Text = List1.Text
       txtCurSong.Text = "1"
       ElseIf txtCount.Text > "1" Then
       txtTime.Text = "0"
       Timer1.Enabled = True
       End If

Check1.Value = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call ReleaseCapture
    Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    frmMain.Left = Me.Left
    frmMain.Top = Me.Top - frmMain.Height
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub List1_DblClick()
pctPlay.Picture = pctOn.Picture
pctCancel.Picture = pctOff.Picture
txtCurSong.Text = List1.ListIndex + 1
frmMain.MediaPlayer1.Stop
frmMain.Text1.Text = List1.Text
frmMain.MediaPlayer1.Filename = List1.Text
frmMain.Command6 = True
End Sub

Private Sub pctBack_DblClick()
pctBack.Left = pctBack.Left + 15
pctBack.Top = pctBack.Top + 15
End Sub

Private Sub pctBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctBack.Left = pctBack.Left + 15
pctBack.Top = pctBack.Top + 15
pctBack.Picture = pctOn.Picture
End Sub

Private Sub pctBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctBack.Left = pctBack.Left - 15
pctBack.Top = pctBack.Top - 15
pctBack.Picture = pctOff.Picture
cmdTrigBack = True
End Sub

Private Sub pctCancel_DblClick()
pctCancel.Left = pctCancel.Left + 15
pctCancel.Top = pctCancel.Top + 15
End Sub

Private Sub pctCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctCancel.Left = pctCancel.Left + 15
pctCancel.Top = pctCancel.Top + 15
End Sub

Private Sub pctCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctCancel.Left = pctCancel.Left - 15
pctCancel.Top = pctCancel.Top - 15
If frmMain.MediaPlayer1.CurrentPosition > 0 Then
pctCancel.Picture = pctOn.Picture
pctPlay.Picture = pctOff.Picture
frmMain.Command5 = True
End If
End Sub

Private Sub pctClear1_DblClick()
pctClear1.Left = pctClear1.Left + 15
pctClear1.Top = pctClear1.Top + 15
End Sub

Private Sub pctClear1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctClear1.Left = pctClear1.Left + 15
pctClear1.Top = pctClear1.Top + 15
pctClear1.Picture = pctOn.Picture
End Sub

Private Sub pctClear1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
       Dim i%
       pctClear1.Left = pctClear1.Left - 15
       pctClear1.Top = pctClear1.Top - 15
       pctClear1.Picture = pctOff.Picture
       If Not txtCount = "0" Then
       List1.RemoveItem List1.ListIndex
       If Text2.Text = "" Then List1.AddItem Text1.Text
       If Text2.Text = "" Then Check1.Value = 1
       txtCount.Text = List1.ListCount
       If txtCount.Text = "1" Then
       txtTime.Text = "0"
       Timer1.Enabled = False
       ElseIf txtCount.Text > "1" Then
       txtTime.Text = "0"
       Timer1.Enabled = True
       End If
       If Not txtCount = "0" Then List1.ListIndex = 0
       txtOut.Text = List1.Text
       txtCurSong.Text = "1"
       Open "c:\Freeplayer.ini" For Output As #1
       For i = 0 To List1.ListCount - 1
       Print #1, List1.List(i)
 
       Next
       Close #1
       If txtCount.Text = "0" Then List1.AddItem "No Entries..."
       txtCurSong.Text = "0"

       End If
End Sub

Private Sub pctClearAll_DblClick()
pctClearAll.Left = pctClearAll.Left + 15
pctClearAll.Top = pctClearAll.Top + 15
End Sub

Private Sub pctClearAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctClearAll.Left = pctClearAll.Left + 15
pctClearAll.Top = pctClearAll.Top + 15
pctClearAll.Picture = pctOn.Picture
End Sub

Private Sub pctClearAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctClearAll.Left = pctClearAll.Left - 15
pctClearAll.Top = pctClearAll.Top - 15
pctClearAll.Picture = pctOff.Picture
List1.Clear
txtCount.Text = "0"
txtOut.Text = ""
txtCurSong.Text = "0"
frmMain.Check1.Value = 0
frmMain.Text1.Text = ""
Open "c:\Freeplayer.ini" For Output As #1
Dim i%
For i = 0 To List1.ListCount - 1
Print #1, ""
 
Next
Close #1
If txtCount.Text = "0" Then
List1.AddItem "No Entries..."
txtCurSong.Text = "0"
End If
End Sub

Private Sub pctEdit_DblClick()
pctEdit.Left = pctEdit.Left + 15
pctEdit.Top = pctEdit.Top + 15
End Sub

Private Sub pctEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctEdit.Left = pctEdit.Left + 15
pctEdit.Top = pctEdit.Top + 15
pctEdit.Picture = pctOn.Picture
End Sub

Private Sub pctEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctEdit.Left = pctEdit.Left - 15
pctEdit.Top = pctEdit.Top - 15
pctEdit.Picture = pctOff.Picture
  'working variables
   Dim c As Integer
   Dim q As Integer
   Dim sFile As String
   Dim startStrg As String
   Dim tmp As String
  
  'dim an array to hold the files selected
   Dim FileArray() As String
   Timer1.Enabled = False
   txtTime.Text = ""
   txtCount.Text = ""
   Check1.Value = 0
  'set the max buffer large enough to retrieve multiple files
   CommonDialog1.DialogTitle = "Load your playlist." & "  Hold the ctrl key to multi-select."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.Filename = ""
   CommonDialog1.Filter = "MP3 Files|*.MP3"
   CommonDialog1.FLAGS = cdlOFNAllowMultiselect Or cdlOFNExplorer
   CommonDialog1.ShowOpen     ' = 1
  
  'assign the string returned from the
  'common dialog to startStrg for further processing.
  'Note that two "nulls" are appended to the
  'end of the string.  This is for use in the StripItem
  'routine below.
   startStrg = CommonDialog1.Filename & Chr(0) & Chr(0)
  
  'Extract each returned filename.
  'If only 1 file was selected, then the string
  'contains the fully-qualified path to the file.
   
  'If more than 1 string file was selected, the
  'string contains the path as the first item,
  'and the FileArray as the rest of the string,
  'each separated by a space.
   For c = 1 To Len(CommonDialog1.Filename)
      
     'extract 1 item from the string
      sFile = StripItem(startStrg, Chr(0))
      
     'if nothing's there, we're done
      If sFile = "" Then Exit For
      
        'redim the filename array
        'to add the new file. FileArray(0) is either the
        'path (if more than 1 file selected), or the
        'fully qualified filename (if only 1 file selected).
         ReDim Preserve FileArray(0 To q)
         FileArray(q) = LCase(sFile)
         
        'increment y by 1 for the next pass
         q = q + 1
      
      Next
      
     'display the results
      Text1.Text = ""
      Text2.Text = ""
      List1.Clear
      
     'starting with 0, and ending with y-1 (because
     'the above will add 1 more 'y'
     'than there is actual array members)
      For c = 0 To q - 1
      
        'if its the first item, display it in Text1,
        'otherwise display the files selected in Text2.
         If c = 0 Then
               Text1.Text = FileArray(c)
         Else: tmp = tmp & Text1.Text & "\" & FileArray(c) & vbCrLf
         List1.AddItem Text1.Text & "\" & FileArray(c)
         End If
      Next
      Text2.Text = tmp
       If Text2.Text = "" Then List1.AddItem Text1.Text
       If Text2.Text = "" Then Check1.Value = 1
       txtCount.Text = List1.ListCount
       If txtCount.Text = "1" Then
       txtTime.Text = "0"
       Timer1.Enabled = False
       ElseIf txtCount.Text > "1" Then
       txtTime.Text = "0"
       Timer1.Enabled = True
       End If
       List1.ListIndex = 0
       txtOut.Text = List1.Text
       txtCurSong.Text = "1"
       If List1.Text = "" Then
       List1.RemoveItem List1.ListIndex
       List1.AddItem "No entries..."
       txtOut.Text = ""
       txtCurSong.Text = "0"
       txtCount.Text = "0"
       ElseIf List1.Text > "" Then
       frmMain.Text1.Text = txtOut.Text
       frmMain.cmdPlaylistT = True
       Open "c:\Freeplayer.ini" For Output As #1
       Dim i%
       For i = 0 To List1.ListCount - 1
       Print #1, List1.List(i)
 
       Next
       Close #1


       End If
End Sub

Private Sub pctNext_DblClick()
pctNext.Left = pctNext.Left + 15
pctNext.Top = pctNext.Top + 15
End Sub

Private Sub pctNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctNext.Left = pctNext.Left + 15
pctNext.Top = pctNext.Top + 15
pctNext.Picture = pctOn.Picture
End Sub

Private Sub pctNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctNext.Left = pctNext.Left - 15
pctNext.Top = pctNext.Top - 15
pctNext.Picture = pctOff.Picture
cmdNext = True
End Sub

Private Sub pctPlay_DblClick()
pctPlay.Left = pctPlay.Left + 15
pctPlay.Top = pctPlay.Top + 15
End Sub

Private Sub pctPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPlay.Left = pctPlay.Left + 15
pctPlay.Top = pctPlay.Top + 15
If txtOut.Text = "" Then
pctPlay.Picture = pctOn.Picture
pctCancel.Picture = pctOff.Picture
End If
End Sub

Private Sub pctPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pctPlay.Left = pctPlay.Left - 15
pctPlay.Top = pctPlay.Top - 15
If txtCount.Text > "" Then frmMain.Check1.Value = 1
frmMain.Text1.Text = txtOut.Text
If txtCount.Text > "" Then frmMain.Command2 = True
If txtOut.Text > "" Then
pctPlay.Picture = pctOn.Picture
pctCancel.Picture = pctOff.Picture
If txtCount.Text > "0" Then
Open "c:\Freeplayer.ini" For Output As #1
Dim i%
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next
Close #1
End If
ElseIf txtCount.Text = "0" Then
pctPlay.Picture = pctOff.Picture
pctCancel.Picture = pctOn.Picture
End If
End Sub

Private Sub Picture1_DblClick()
Picture1.Left = Picture1.Left + 15
Picture1.Top = Picture1.Top + 15
If Picture1.Picture = pctOff.Picture Then
Picture1.Picture = pctOn.Picture
ElseIf Picture1.Picture = pctOn.Picture Then
Picture1.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Left = Picture1.Left + 15
Picture1.Top = Picture1.Top + 15
If Picture1.Picture = pctOff.Picture Then
Picture1.Picture = pctOn.Picture
ElseIf Picture1.Picture = pctOn.Picture Then
Picture1.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Left = Picture1.Left - 15
Picture1.Top = Picture1.Top - 15
End Sub

Private Sub Picture10_DblClick()
Picture10.Left = Picture10.Left + 15
Picture10.Top = Picture10.Top + 15
If Picture10.Picture = pctOff.Picture Then
Picture10.Picture = pctOn.Picture
ElseIf Picture10.Picture = pctOn.Picture Then
Picture10.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture10.Left = Picture10.Left + 15
Picture10.Top = Picture10.Top + 15
If Picture10.Picture = pctOff.Picture Then
Picture10.Picture = pctOn.Picture
ElseIf Picture10.Picture = pctOn.Picture Then
Picture10.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture10.Left = Picture10.Left - 15
Picture10.Top = Picture10.Top - 15
End Sub

Private Sub Picture11_DblClick()
Picture11.Left = Picture11.Left + 15
Picture11.Top = Picture11.Top + 15
If Picture11.Picture = pctOff.Picture Then
Picture11.Picture = pctOn.Picture
ElseIf Picture11.Picture = pctOn.Picture Then
Picture11.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture11.Left = Picture11.Left + 15
Picture11.Top = Picture11.Top + 15
If Picture11.Picture = pctOff.Picture Then
Picture11.Picture = pctOn.Picture
ElseIf Picture11.Picture = pctOn.Picture Then
Picture11.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture11.Left = Picture11.Left - 15
Picture11.Top = Picture11.Top - 15
End Sub

Private Sub Picture8_DblClick()
Picture8.Left = Picture8.Left + 15
Picture8.Top = Picture8.Top + 15
If Picture8.Picture = pctOff.Picture Then
Picture8.Picture = pctOn.Picture
ElseIf Picture8.Picture = pctOn.Picture Then
Picture8.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture8.Left = Picture8.Left + 15
Picture8.Top = Picture8.Top + 15
If Picture8.Picture = pctOff.Picture Then
Picture8.Picture = pctOn.Picture
ElseIf Picture8.Picture = pctOn.Picture Then
Picture8.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture8.Left = Picture8.Left - 15
Picture8.Top = Picture8.Top - 15
End Sub

Private Sub Picture9_DblClick()
Picture9.Left = Picture9.Left + 15
Picture9.Top = Picture9.Top + 15
If Picture9.Picture = pctOff.Picture Then
Picture9.Picture = pctOn.Picture
ElseIf Picture9.Picture = pctOn.Picture Then
Picture9.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture9.Left = Picture9.Left + 15
Picture9.Top = Picture9.Top + 15
If Picture9.Picture = pctOff.Picture Then
Picture9.Picture = pctOn.Picture
ElseIf Picture9.Picture = pctOn.Picture Then
Picture9.Picture = pctOff.Picture
End If
End Sub

Private Sub Picture9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture9.Left = Picture9.Left - 15
Picture9.Top = Picture9.Top - 15
End Sub

Private Sub Text2_GotFocus()

     Text2.SelStart = 0
     Text2.SelLength = 100
End Sub
