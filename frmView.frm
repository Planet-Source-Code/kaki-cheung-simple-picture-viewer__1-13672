VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmView 
   BackColor       =   &H00000000&
   Caption         =   "Viewer"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5790
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlView 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "(Default type)"
   End
   Begin MSComctlLib.ProgressBar pbrView 
      Height          =   195
      Left            =   3660
      TabIndex        =   2
      Top             =   5580
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sbrView 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5535
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3810
            MinWidth        =   3810
            Text            =   "Dimension"
            TextSave        =   "Dimension"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Text            =   "Size"
            TextSave        =   "Size"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4101
            Picture         =   "frmView.frx":0442
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picView 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   3135
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuDump 
         Caption         =   "&Dump it"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPics 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private totalPic As Integer
Private lastIndex As Integer
Private StartMove As Integer
Private OldX As Single, OldY As Single

Private Sub ViewPic(fName As String, Index As Integer)
On Error GoTo errHandle
    Load picView(Index)
    picView(Index).Tag = fName
    picView(Index) = LoadPicture(fName)
    picView(Index).Visible = True
    Load mnuPics(Index)
    mnuPics(Index).Caption = Right(fName, Len(fName) - InStrRev(fName, "\"))
    mnuPics(Index).Visible = True
    Exit Sub
    
errHandle:
    MsgBox Err.Description + ": " + fName, vbCritical
    If picView(Index).Picture = 0 Then Unload picView(Index)
    Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
    If App.PrevInstance = True Then End
    If Command = "" Then Exit Sub
    Dim files() As String, pos As Integer
    files = Split(Command)
    For pos = 0 To UBound(files)
        ViewPic files(pos), pos + 1
    Next
    totalPic = UBound(files) + 1
    If Err Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And mnuPics.Count > 1 Then
        mnuSave.Visible = False
        mnuDump.Visible = False
        mnuSep.Visible = False
        PopupMenu mnuPopup
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    pbrView.Visible = True
    pbrView.Max = Data.files.Count
    
    Dim i As Integer
    For i = 1 To Data.files.Count
        ViewPic Data.files(i), totalPic + i
        pbrView.Value = i
    Next
    totalPic = totalPic + Data.files.Count
    pbrView.Visible = False
    pbrView.Value = 0
    
If Err Then MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        pbrView.Move 3660, ScaleHeight - 210
        pbrView.Width = ScaleWidth - 3660 - 315
    End If
End Sub

Private Sub mnuDump_Click()
    Unload picView(lastIndex)
    Unload mnuPics(lastIndex)
    lastIndex = 0
    If picView.Count = 1 Then
        Caption = "Viewer"
        sbrView.Panels(1).Text = "Dimension"
        sbrView.Panels(2).Text = "Size"
    End If
End Sub

Private Sub mnuPics_Click(Index As Integer)
    picView(Index).ZOrder
    sbrView.ZOrder
    pbrView.ZOrder
    mnuPics(Index).Checked = True
    mnuPics(lastIndex).Checked = False
    lastIndex = Index
End Sub

Private Sub mnuSave_Click()
On Error GoTo errHandle
    cdlView.Flags = cdlOFNCreatePrompt + cdlOFNExplorer + cdlOFNHideReadOnly + _
        cdlOFNLongNames + cdlOFNOverwritePrompt
    cdlView.FileName = Right(picView(lastIndex).Tag, Len(picView(lastIndex).Tag) - InStrRev(picView(lastIndex).Tag, "\"))
    cdlView.ShowSave
    SavePicture picView(lastIndex), cdlView.FileName
    Exit Sub
    
errHandle:
    If Err <> 32755 Then MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub picView_GotFocus(Index As Integer)
    picView(Index).ZOrder
    sbrView.ZOrder
    pbrView.ZOrder
    Caption = picView(Index).Tag
    sbrView.Panels(1).Text = "Dimension: " & picView(Index).Width \ 16 & " x " & picView(Index).Height \ 16
On Error Resume Next
    Dim SizePic As Long, size As Long, unit As String
    SizePic = FileLen(picView(Index).Tag)
    size = IIf(SizePic \ 1024 > 1, SizePic \ 1024, SizePic)
    unit = IIf(SizePic \ 1024 > 1, " KB", " Bytes")
    sbrView.Panels(2).Text = "Size: " & Format(size, "#,##0") & unit
End Sub

Private Sub picView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mnuSave.Visible = True
        mnuDump.Visible = True
        mnuSep.Visible = True
        picView(Index).ZOrder
        sbrView.ZOrder
        pbrView.ZOrder
        mnuPics(Index).Checked = True
        mnuPics(lastIndex).Checked = False
        lastIndex = Index
        PopupMenu mnuPopup
    ElseIf Button = vbLeftButton Then
        StartMove = True
        OldX = X
        OldY = Y
    End If
End Sub

Private Sub picView_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If StartMove And Button = vbLeftButton Then
        picView(Index).Move picView(Index).Left + (X - OldX), picView(Index).Top + (Y - OldY)
    End If
End Sub

Private Sub picView_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartMove = False
End Sub

Private Sub picView_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
