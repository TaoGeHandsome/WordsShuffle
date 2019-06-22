VERSION 5.00
Begin VB.Form FormShuffle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Words Shuffle by TaoGe"
   ClientHeight    =   2715
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8055
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   8055
   StartUpPosition =   1  '所有者中心
   Begin VB.Label lbl_unit 
      Caption         =   "Unit 1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label lbl_notice_chn 
      Caption         =   "Chinese: ON"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbl_notice_jpn 
      Caption         =   "Japanese: OFF"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6720
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbl_notice_hira 
      Caption         =   "Hiragana: OFF"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbl_notice 
      Alignment       =   2  'Center
      Caption         =   "Press Space to Shuffle"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label lbl_hira 
      Alignment       =   2  'Center
      Caption         =   "Hira"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lbl_jpn 
      Alignment       =   2  'Center
      Caption         =   "JPN"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lbl_chn 
      Alignment       =   2  'Center
      Caption         =   "CHN"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Menu mn_file 
      Caption         =   "&File"
      Begin VB.Menu mn_db 
         Caption         =   "&Databases"
         Begin VB.Menu mn_dbs 
            Caption         =   "Database"
            Index           =   0
         End
         Begin VB.Menu line1 
            Caption         =   "-"
         End
         Begin VB.Menu mn_load 
            Caption         =   "&Load..."
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mn_editor 
         Caption         =   "&Editor..."
         Shortcut        =   ^E
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mn_unit 
      Caption         =   "&Units"
      Begin VB.Menu mn_units 
         Caption         =   "Unit"
         Index           =   0
      End
   End
   Begin VB.Menu mn_about 
      Caption         =   "&About"
      Begin VB.Menu mn_help 
         Caption         =   "&Help..."
         Shortcut        =   ^H
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_info 
         Caption         =   "&Info..."
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "FormShuffle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
     hWndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags    As Long
     lpfnCallback     As Long
     lParam     As Long
     iImage     As Long
End Type
Dim HIRA_ON As Boolean, JPN_ON As Boolean, CHN_ON As Boolean
Dim VIEW_ENABLED As Boolean
Dim CL_ENABLED As OLE_COLOR, CL_DISABLED As OLE_COLOR
Dim curpos As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And VIEW_ENABLED = True Then
        curpos = curpos + 1
        If curpos > UBound(wlist) Then
            shuffle_words
            curpos = 0
        End If
        set_show_word curpos
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 72 Or KeyAscii = 104 Then
        HIRA_ON = IIf(HIRA_ON, False, True)
    ElseIf KeyAscii = 74 Or KeyAscii = 106 Then
        JPN_ON = IIf(JPN_ON, False, True)
    ElseIf KeyAscii = 67 Or KeyAscii = 99 Then
        CHN_ON = IIf(CHN_ON, False, True)
    End If
    set_view
End Sub

Private Sub Form_Load()
    Randomize
    'Color const
    CL_ENABLED = &H4000&
    CL_DISABLED = &H808080
    
    'Range position
    lbl_chn.Width = FormShuffle.Width * 0.9
    lbl_hira.Width = FormShuffle.Width * 0.9
    lbl_jpn.Width = FormShuffle.Width * 0.9
    lbl_chn.Left = (FormShuffle.Width - lbl_chn.Width) / 2
    lbl_hira.Left = (FormShuffle.Width - lbl_hira.Width) / 2
    lbl_jpn.Left = (FormShuffle.Width - lbl_jpn.Width) / 2
    
    Dim try_path As String
    try_path = search_first_valid_db(App.Path & "\")
    MsgBox try_path
    load_db_form App.Path & "\" & try_path
End Sub

Private Sub mn_dbs_Click(Index As Integer)
    load_db_form App.Path & "\" & mn_dbs(Index).Caption
End Sub

Private Sub mn_editor_Click()
    Me.Hide
    FormEditor.Show
End Sub

Private Sub mn_load_Click()
     Dim lpIDList As Long
     Dim sBuffer As String
     Dim szTitle As String
     Dim tBrowseInfo As BrowseInfo
     szTitle = App.Path
     With tBrowseInfo
          .hWndOwner = Me.hWnd
          .lpszTitle = lstrcat(szTitle, "")
          .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
     End With

     lpIDList = SHBrowseForFolder(tBrowseInfo)
     If (lpIDList) Then
          sBuffer = Space(MAX_PATH)
          SHGetPathFromIDList lpIDList, sBuffer
          sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
          load_db_form sBuffer
     End If
End Sub

Private Sub mn_units_Click(Index As Integer)
    If load_unit(Index) Then
        lbl_unit.Caption = ulist(Index).uname
        init_view True
        set_show_word 0
    End If
End Sub

Private Sub mn_exit_Click()
    End
End Sub

Private Sub rm_unit_menu()
    Dim i As Integer
    mn_units(0).Caption = "(Units)"
    mn_units(0).Enabled = False
    For i = mn_units.LBound + 1 To mn_units.UBound
        Unload mn_units(i)
    Next i
End Sub

Private Sub add_unit_menu()
    rm_unit_menu
    Dim i As Integer
    For i = LBound(ulist) To UBound(ulist)
        If i > 0 Then Load mn_units(i)
        mn_units(i).Caption = ulist(i).uname
        mn_units(i).Enabled = True
    Next i
End Sub

Private Sub rm_db_menu()
    Dim i As Integer
    mn_dbs(0).Caption = "(Databases Found)"
    mn_dbs(0).Enabled = False
    For i = mn_dbs.LBound + 1 To mn_dbs.UBound
        Unload mn_dbs(i)
    Next i
End Sub

Private Sub add_db_menu(fpath As String)
    If mn_dbs(0).Enabled = True Then Load mn_dbs(mn_dbs.UBound + 1)
    mn_dbs(mn_dbs.UBound).Caption = fpath
    mn_dbs(mn_dbs.UBound).Enabled = True
End Sub

Private Function search_first_valid_db(filepath As String) As String
    rm_db_menu
    search_first_valid_db = ""
    Dim folders() As String, cnt As Integer, f As String, i As Integer
    cnt = 0
    f = Dir(filepath, vbDirectory)
    ReDim folders(0)
    folders(0) = f
    Do While f <> ""
        f = Dir
        cnt = cnt + 1
        ReDim Preserve folders(cnt)
        folders(cnt) = f
    Loop
    For i = 0 To cnt - 1
        If is_valid_db(folders(i)) Then
            If search_first_valid_db = "" Then search_first_valid_db = folders(i)
            add_db_menu folders(i)
        End If
    Next i
End Function

Public Sub load_db_form(fpath As String)
    If load_database(fpath) Then
        add_unit_menu
        no_unit_view
        mn_unit.Enabled = True
        Dim res, fname As String
        res = Split(fpath, "\")
        fname = res(UBound(res))
        FormShuffle.Caption = "Words Shuffle by TaoGe  [" & fname & "]"
    Else
        rm_unit_menu
        no_db_view
        mn_unit.Enabled = False
        MsgBox "Not a valid database!", vbInformation, "Error"
        FormShuffle.Caption = "Words Shuffle by TaoGe"
    End If
End Sub


Public Sub no_db_view()
    CHN_ON = False
    HIRA_ON = False
    JPN_ON = False
    VIEW_ENABLED = False
    set_view
    mn_editor.Enabled = False
    lbl_unit.Caption = "No units found."
    lbl_notice.Caption = "Click File->Load to load"
End Sub

Public Sub no_unit_view()
    CHN_ON = False
    HIRA_ON = False
    JPN_ON = False
    VIEW_ENABLED = False
    set_view
    mn_editor.Enabled = True
    lbl_unit.Caption = "Choose a unit."
    lbl_notice.Caption = "Click Unit to choose"
End Sub

Public Sub init_view(is_shuffle As Boolean)
    CHN_ON = True
    HIRA_ON = False
    JPN_ON = False
    VIEW_ENABLED = True
    lbl_notice.Caption = "Press Space to Shuffle"
    set_view
    curpos = 0
    If is_shuffle Then shuffle_words
End Sub

Public Sub set_view()
    lbl_notice_hira.Caption = "Hiragana: " + IIf(HIRA_ON And VIEW_ENABLED, "ON", "OFF")
    lbl_notice_chn.Caption = "Chinese: " + IIf(CHN_ON And VIEW_ENABLED, "ON", "OFF")
    lbl_notice_jpn.Caption = "Japanese: " + IIf(JPN_ON And VIEW_ENABLED, "ON", "OFF")
    lbl_notice_chn.ForeColor = IIf(CHN_ON And VIEW_ENABLED, CL_ENABLED, CL_DISABLED)
    lbl_notice_hira.ForeColor = IIf(HIRA_ON And VIEW_ENABLED, CL_ENABLED, CL_DISABLED)
    lbl_notice_jpn.ForeColor = IIf(JPN_ON And VIEW_ENABLED, CL_ENABLED, CL_DISABLED)
    lbl_jpn.Visible = JPN_ON And VIEW_ENABLED
    lbl_hira.Visible = HIRA_ON And VIEW_ENABLED
    lbl_chn.Visible = CHN_ON And VIEW_ENABLED
End Sub

Private Sub set_show_word(Index As Integer)
    lbl_chn.Caption = TYPE_DESCRIBE(wlist(Index).wtype) & wlist(Index).chn
    lbl_hira.Caption = wlist(Index).hira
    lbl_jpn.Caption = wlist(Index).jpn
End Sub

