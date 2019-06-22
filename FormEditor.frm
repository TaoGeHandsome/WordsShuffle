VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FormEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Words Shuffle - Editor By Taoge"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10710
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frame_preview 
      Caption         =   "[Preview]"
      Height          =   2175
      Left            =   1320
      TabIndex        =   32
      Top             =   3960
      Width           =   8145
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
         Left            =   3120
         TabIndex        =   35
         Top             =   240
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
         Left            =   3120
         TabIndex        =   34
         Top             =   1560
         Width           =   4335
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
         Left            =   3120
         TabIndex        =   33
         Top             =   1080
         Width           =   4335
      End
   End
   Begin RichTextLib.RichTextBox txt_title 
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"FormEditor.frx":0000
   End
   Begin VB.Frame frame_word_list 
      Caption         =   "[Word List]"
      Height          =   3495
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   8295
      Begin VB.CommandButton btn_down 
         Caption         =   "Down"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton btn_up 
         Caption         =   "Up"
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   2760
         Width           =   615
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Proper Noun"
         Height          =   255
         Index           =   13
         Left            =   6480
         TabIndex        =   31
         Top             =   3120
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Interjection"
         Height          =   255
         Index           =   12
         Left            =   6480
         TabIndex        =   30
         Top             =   2880
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Conjunction"
         Height          =   255
         Index           =   11
         Left            =   6480
         TabIndex        =   29
         Top             =   2640
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Word Piece"
         Height          =   255
         Index           =   10
         Left            =   6480
         TabIndex        =   28
         Top             =   2400
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Adverb"
         Height          =   255
         Index           =   9
         Left            =   6480
         TabIndex        =   27
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Adj II"
         Height          =   255
         Index           =   8
         Left            =   7320
         TabIndex        =   26
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Adj I"
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   25
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Verb III"
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Verb II"
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Verb I"
         Height          =   255
         Index           =   4
         Left            =   6480
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Interrogative"
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Pronoun"
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Noun"
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton opt_wtype 
         Caption         =   "Unknown or NULL"
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox list_words 
         Height          =   3120
         ItemData        =   "FormEditor.frx":009D
         Left            =   120
         List            =   "FormEditor.frx":009F
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox txt_chn 
         Height          =   405
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   714
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         TextRTF         =   $"FormEditor.frx":00A1
      End
      Begin RichTextLib.RichTextBox txt_hira 
         Height          =   405
         Left            =   2280
         TabIndex        =   9
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   714
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"FormEditor.frx":013E
      End
      Begin RichTextLib.RichTextBox txt_jpn 
         Height          =   405
         Left            =   2280
         TabIndex        =   10
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   714
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"FormEditor.frx":01DB
      End
      Begin VB.CommandButton btn_add_word 
         Caption         =   "Add Word"
         Height          =   615
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton btn_remove_word 
         Caption         =   "Remove Word"
         Height          =   615
         Left            =   5040
         MaskColor       =   &H80000010&
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbl_wtype 
         Caption         =   "Part of speech:"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_info_chn 
         Caption         =   "Chinese:"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl_info_hira 
         Caption         =   "Hiragana: "
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_info_jpn 
         Caption         =   "Japanese:"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.CommandButton btn_save 
      Caption         =   "Save"
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin RichTextLib.RichTextBox txt_filename 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"FormEditor.frx":0278
   End
   Begin VB.ListBox list_files 
      Height          =   3660
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbl_title 
      Caption         =   "Title:"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lbl_filename 
      Caption         =   "File name:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lbl_files 
      Caption         =   "Data files:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu mn_editor_file 
      Caption         =   "&File"
      Begin VB.Menu mn_editor_save 
         Caption         =   "&Save Word List"
         Shortcut        =   ^S
      End
      Begin VB.Menu mn_editor_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_editor_add_file 
         Caption         =   "&Add Data File"
         Shortcut        =   ^N
      End
      Begin VB.Menu mn_editor_rm_file 
         Caption         =   "&Remove Data File"
         Shortcut        =   ^R
      End
      Begin VB.Menu mn_editor_shift_up 
         Caption         =   "Shift &Up Data File"
         Shortcut        =   ^J
      End
      Begin VB.Menu mn_editor_shift_down 
         Caption         =   "Shift &Down Data File"
         Shortcut        =   ^K
      End
      Begin VB.Menu mn_editor_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_editor_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mn_editor_about 
      Caption         =   "&About"
      Begin VB.Menu mn_editor_help 
         Caption         =   "&Help..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mn_editor_line3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_editor_info 
         Caption         =   "&Info"
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "FormEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim is_file_changed As Boolean, cur_idx As Integer, ctrlFlag As Boolean

Private Sub btn_add_word_Click()
    set_modified_status
    If list_words.ListCount = 0 Then
        ReDim wlist(0)
    Else
        ReDim Preserve wlist(UBound(wlist) + 1)
    End If
    wlist(UBound(wlist)).chn = "(New word)"
    wlist(UBound(wlist)).hira = ""
    wlist(UBound(wlist)).jpn = ""
    list_words.AddItem wlist(UBound(wlist)).chn
    list_words.ListIndex = UBound(wlist)
    update_input
    set_selection txt_chn
End Sub

Private Sub btn_down_Click()
    If list_words.ListIndex = list_words.ListCount - 1 Then Exit Sub
    set_modified_status
    Dim i As Integer, tw As Word, ts As String
    i = list_words.ListIndex
    tw = wlist(i)
    wlist(i) = wlist(i + 1)
    wlist(i + 1) = tw
    ts = list_words.List(i)
    list_words.List(i) = list_words.List(i + 1)
    list_words.List(i + 1) = ts
    list_words.ListIndex = i + 1
End Sub

Private Sub btn_remove_word_Click()
    If list_words.ListIndex < 0 Then Exit Sub
    set_modified_status
    Dim i As Integer, idx As Integer
    idx = list_words.ListIndex
    For i = idx To UBound(wlist) - 1
        wlist(i) = wlist(i + 1)
    Next i
    If UBound(wlist) > 0 Then
        ReDim Preserve wlist(UBound(wlist) - 1)
    Else
        Erase wlist()
    End If
    list_words.RemoveItem idx
    If idx = list_words.ListCount And idx > 0 Then
        list_words.ListIndex = idx - 1
    Else
        list_words.ListIndex = IIf(list_words.ListCount = 0, -1, idx)
    End If
    If list_words.ListCount = 0 Then
        clear_input
        disable_modify
    End If
End Sub

Private Sub btn_save_Click()
    save_title_info cur_idx, txt_filename.Text, txt_title.Text
    list_files.Clear
    Dim i As Integer
    For i = 0 To UBound(data_file_name)
        list_files.AddItem data_file_name(i)
    Next i
End Sub

Private Sub btn_up_Click()
    If list_words.ListIndex < 1 Then Exit Sub
    set_modified_status
    Dim i As Integer, tw As Word, ts As String
    i = list_words.ListIndex
    tw = wlist(i)
    wlist(i) = wlist(i - 1)
    wlist(i - 1) = tw
    ts = list_words.List(i)
    list_words.List(i) = list_words.List(i - 1)
    list_words.List(i - 1) = ts
    list_words.ListIndex = i - 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then ctrlFlag = True
    If KeyCode = 68 And ctrlFlag Then        ' Ctrl + D
        btn_remove_word_Click
        KeyCode = 0
    End If
    If KeyCode = 73 And ctrlFlag Then        ' Ctrl + I
        insert_word
        KeyCode = 0
    End If
    If KeyCode = 79 And ctrlFlag Then        ' Ctrl + O
        btn_up_Click
        KeyCode = 0
    End If
    If KeyCode = 80 And ctrlFlag Then        ' Ctrl + P
        btn_down_Click
        KeyCode = 0
    End If
    If KeyCode = 112 Then       ' F1
        If list_words.ListIndex > 0 Then list_words.ListIndex = list_words.ListIndex - 1
        KeyCode = 0
    End If
    If KeyCode = 113 Then       ' F2
        If list_words.ListIndex > -1 And list_words.ListIndex <> list_words.ListCount - 1 Then list_words.ListIndex = list_words.ListIndex + 1
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then ctrlFlag = False
End Sub

Private Sub Form_Load()
    load_database DATABASE
    add_unit_list
    is_file_changed = False
    mn_editor_save.Enabled = False
    btn_save.Enabled = False
    txt_filename.Enabled = False
    txt_title.Enabled = False
    btn_remove_word.Enabled = False
    btn_add_word.Enabled = False
    btn_up.Enabled = False
    btn_down.Enabled = False
    disable_modify
    
    lbl_chn.Caption = ""
    lbl_hira.Caption = ""
    lbl_jpn.Caption = ""
    
    'Range position
    lbl_chn.Width = frame_preview.Width * 0.9
    lbl_hira.Width = frame_preview.Width * 0.9
    lbl_jpn.Width = frame_preview.Width * 0.9
    lbl_chn.Left = (FormShuffle.Width - lbl_chn.Width) / 2
    lbl_hira.Left = (FormShuffle.Width - lbl_hira.Width) / 2
    lbl_jpn.Left = (FormShuffle.Width - lbl_jpn.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim result
    If is_file_changed Then
        result = MsgBox("The file has been modified." & vbCrLf & "Do you want to save it?", vbYesNo, "Notice")
        If result = vbYes Then mn_editor_save_Click
    End If
    FormShuffle.load_db_form DATABASE
    FormShuffle.Show
    'End
End Sub

Private Sub add_unit_list()
    list_files.Clear
    Dim i As Integer, fname As String, items
    ReDim data_file_name(0)
    For i = LBound(ulist) To UBound(ulist)
        items = Split(ulist(i).fpath, "\")
        fname = items(UBound(items))
        ReDim Preserve data_file_name(i)
        data_file_name(i) = Left(fname, Len(fname) - 5)
        list_files.AddItem data_file_name(i)
    Next i
End Sub

Private Sub list_files_Click()
    Dim result
    If is_file_changed Then
        result = MsgBox("The file has been modified." & vbCrLf & "Do you want to save it?", vbYesNo, "Notice")
        If result = vbYes Then mn_editor_save_Click
        remove_modified_status
    End If
    cur_idx = list_files.ListIndex
    If Not load_unit(cur_idx) Then Exit Sub
    txt_filename.Text = list_files.List(cur_idx)
    txt_title.Text = ulist(cur_idx).uname
    Dim i As Integer
    list_words.Clear
    If wlist(0).chn <> "" Then
        For i = LBound(wlist) To UBound(wlist)
            list_words.AddItem wlist(i).chn
        Next i
    End If
    txt_filename.Enabled = True
    txt_title.Enabled = True
    btn_add_word.Enabled = True
    btn_save.Enabled = False
    clear_input
    disable_modify
End Sub

Private Sub list_words_Click()
    If txt_chn.Enabled = False Then enable_modify
    'If list_words.ListIndex = list_words.ListCount - 1 Then btn_add_word_Click
    update_input
End Sub

Private Sub mn_editor_add_file_Click()
    Dim i As Integer, idx As Integer
    ReDim Preserve ulist(UBound(ulist) + 1)
    ReDim Preserve data_file_name(UBound(data_file_name) + 1)
    data_file_name(UBound(data_file_name)) = "(new_file)"
    
    ulist(UBound(ulist)).fpath = DATABASE & "\(new_file).data"
    ulist(UBound(ulist)).uname = "new title"
    touch_file ulist(UBound(ulist)).fpath
    list_files.AddItem "(new_file)"
    list_files.ListIndex = UBound(ulist)
    ' save_unit_list
End Sub

Private Sub mn_editor_exit_Click()
    'End
    Unload Me
End Sub

Private Sub mn_editor_rm_file_Click()
    Dim i As Integer, idx As Integer
    idx = list_files.ListIndex
    If idx < 0 Then Exit Sub
    If list_files.ListCount = 1 Then
        MsgBox "A database must have at least one data file!", vbOKOnly, "Error"
        Exit Sub
    End If
    If MsgBox("Do You want to remove [" & list_files.List(idx) & "] file?" & vbCrLf _
        & "This operation CAN NOT BE UNDO!", vbYesNo, "Warning") = vbNo Then Exit Sub
    ' Remove
    On Error GoTo no_rm_unit
    If Dir(ulist(idx).fpath) = "" Then
        MsgBox "Failed to remove!", vbOKOnly, "Error"
        Exit Sub
    End If
    Kill ulist(idx).fpath
    For i = idx To UBound(ulist) - 1
        ulist(i) = ulist(i + 1)
        data_file_name(i) = data_file_name(i + 1)
    Next i
    
    ReDim Preserve data_file_name(UBound(ulist) - 1)
    ReDim Preserve ulist(UBound(ulist) - 1)
    
    list_files.RemoveItem idx
    If idx = list_files.ListCount And idx > 0 Then
        list_files.ListIndex = idx - 1
    Else
        list_files.ListIndex = IIf(list_files.ListCount = 0, -1, idx)
    End If
no_rm_unit:
    save_unit_list
End Sub

Private Sub mn_editor_save_Click()
    save_word_list cur_idx
    remove_modified_status
End Sub

Private Sub mn_editor_shift_down_Click()
    If list_files.ListIndex < 0 Or cur_idx = list_files.ListCount - 1 Then Exit Sub
    Dim i As Integer, ts As String
    i = list_files.ListIndex
    swap_unit ulist(i), ulist(i + 1)
    swap_str data_file_name(i), data_file_name(i + 1)
    ts = list_files.List(i)
    list_files.List(i) = list_files.List(i + 1)
    list_files.List(i + 1) = ts
    list_files.ListIndex = i + 1
    save_unit_list
End Sub

Private Sub mn_editor_shift_up_Click()
    If cur_idx < 1 Then Exit Sub
    Dim i As Integer, ts As String
    i = list_files.ListIndex
    swap_unit ulist(i), ulist(i - 1)
    swap_str data_file_name(i), data_file_name(i - 1)
    ts = list_files.List(i)
    list_files.List(i) = list_files.List(i - 1)
    list_files.List(i - 1) = ts
    list_files.ListIndex = i - 1
    save_unit_list
End Sub

Private Sub opt_wtype_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    update_view
    If Index <> wlist(list_words.ListIndex).wtype Then modify_wlist
End Sub

Private Sub txt_chn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then set_selection txt_hira
    If KeyCode = 9 Then
        If Shift = 0 And list_words.ListIndex < list_words.ListCount - 1 Then
            list_words.ListIndex = list_words.ListIndex + 1
            update_input
        ElseIf Shift = 1 Then
            list_words.ListIndex = IIf(list_words.ListIndex > 0, list_words.ListIndex - 1, _
                                        list_words.ListCount - 1)
            update_input
        Else
            btn_add_word_Click
        End If
        set_selection txt_chn
        KeyCode = 0
    End If
End Sub

Private Sub txt_hira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then set_selection txt_jpn
    If KeyCode = 9 Then
        If Shift = 0 And list_words.ListIndex < list_words.ListCount - 1 Then
            list_words.ListIndex = list_words.ListIndex + 1
            update_input
        ElseIf Shift = 1 Then
            list_words.ListIndex = IIf(list_words.ListIndex > 0, list_words.ListIndex - 1, _
                                        list_words.ListCount - 1)
            update_input
        Else
            btn_add_word_Click
        End If
        set_selection txt_hira
        KeyCode = 0
    End If
End Sub

Private Sub txt_jpn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If list_words.ListIndex < list_words.ListCount - 1 Then
            list_words.ListIndex = list_words.ListIndex + 1
            update_input
        Else
            btn_add_word_Click
        End If
        set_selection txt_chn
    End If
    If KeyCode = 9 Then
        If Shift = 0 And list_words.ListIndex < list_words.ListCount - 1 Then
            list_words.ListIndex = list_words.ListIndex + 1
            update_input
        ElseIf Shift = 1 Then
            list_words.ListIndex = IIf(list_words.ListIndex > 0, list_words.ListIndex - 1, _
                                        list_words.ListCount - 1)
            update_input
        Else
            btn_add_word_Click
        End If
        set_selection txt_jpn
        KeyCode = 0
    End If
End Sub

Private Sub txt_filename_LostFocus()
    txt_filename.Text = Trim(txt_filename.Text)
End Sub

Private Sub txt_filename_Change()
    If Trim(txt_filename.Text) <> data_file_name(cur_idx) And Not is_file_changed Then
        btn_save.Enabled = True
    Else
        btn_save.Enabled = False
    End If
End Sub

Private Sub txt_title_LostFocus()
    txt_title.Text = Trim(txt_title.Text)
End Sub

Private Sub txt_title_Change()
    If Trim(txt_title.Text) <> ulist(cur_idx).uname And Not is_file_changed Then
        btn_save.Enabled = True
    Else
        btn_save.Enabled = False
    End If
End Sub

Private Sub txt_chn_Change()
    If list_words.ListIndex < 0 Then Exit Sub
    update_view
    If txt_chn.Text <> wlist(list_words.ListIndex).chn Then modify_wlist
End Sub

Private Sub txt_hira_Change()
    If list_words.ListIndex < 0 Then Exit Sub
    update_view
    If txt_hira.Text <> wlist(list_words.ListIndex).hira Then modify_wlist
End Sub

Private Sub txt_jpn_Change()
    If list_words.ListIndex < 0 Then Exit Sub
    update_view
    If txt_jpn.Text <> wlist(list_words.ListIndex).jpn Then modify_wlist
End Sub

Private Sub disable_modify()
    txt_chn.Enabled = False
    txt_hira.Enabled = False
    txt_jpn.Enabled = False
    btn_remove_word.Enabled = False
    btn_up.Enabled = False
    btn_down.Enabled = False
    Dim i As Integer
    For i = 0 To opt_wtype.UBound
        opt_wtype(i).Enabled = False
    Next i
End Sub

Private Sub enable_modify()
    txt_chn.Enabled = True
    txt_hira.Enabled = True
    txt_jpn.Enabled = True
    btn_remove_word.Enabled = True
    btn_up.Enabled = True
    btn_down.Enabled = True
    Dim i As Integer
    For i = 0 To opt_wtype.UBound
        opt_wtype(i).Enabled = True
    Next i
End Sub

Private Sub clear_input()
    Dim i As Integer
    For i = 0 To opt_wtype.UBound
        opt_wtype(i).Value = False
    Next i
    txt_chn.Text = ""
    txt_hira.Text = ""
    txt_jpn.Text = ""
    update_view
End Sub

Private Sub update_input()
    Dim i As Integer
    txt_chn.Text = wlist(list_words.ListIndex).chn
    txt_hira.Text = wlist(list_words.ListIndex).hira
    txt_jpn.Text = wlist(list_words.ListIndex).jpn
    For i = 0 To opt_wtype.UBound
        opt_wtype(i).Value = False
    Next i
    opt_wtype(wlist(list_words.ListIndex).wtype).Value = True
    update_view
End Sub

Private Sub insert_word()
    Dim i As Integer, ins As Integer
    ins = list_words.ListIndex
    If ins = -1 Then btn_add_word_Click: Exit Sub
    
    set_modified_status
    If list_words.ListIndex = 0 Then
        ReDim wlist(0)
    Else
        ReDim Preserve wlist(UBound(wlist) + 1)
    End If
    list_words.AddItem ""
    
    For i = UBound(wlist) To ins Step -1
        wlist(i) = wlist(i - 1)
        list_words.List(i) = list_words.List(i - 1)
    Next i
    wlist(ins).chn = "(New word)"
    wlist(ins).hira = ""
    wlist(ins).jpn = ""
    list_words.ListIndex = ins
    update_input
    set_selection txt_chn
End Sub

Private Sub update_view()
    Dim i As Integer, append As String
    For i = 0 To opt_wtype.UBound
        If opt_wtype(i).Value Then append = TYPE_DESCRIBE(i)
    Next i
    If (list_words.ListIndex <> -1) Then list_words.List(list_words.ListIndex) = txt_chn.Text
    lbl_chn.Caption = append & txt_chn.Text
    lbl_hira.Caption = txt_hira.Text
    lbl_jpn.Caption = txt_jpn.Text
End Sub

Private Sub set_modified_status()
    is_file_changed = True
    btn_save.Enabled = False
    mn_editor_save.Enabled = True
    FormEditor.Caption = "Words Shuffle - Editor By Taoge [Modified]"
End Sub

Private Sub remove_modified_status()
    is_file_changed = False
    mn_editor_save.Enabled = False
    FormEditor.Caption = "Words Shuffle - Editor By Taoge"
End Sub

Private Sub modify_wlist()
    set_modified_status
    Dim i As Integer
    wlist(list_words.ListIndex).chn = txt_chn.Text
    wlist(list_words.ListIndex).hira = txt_hira.Text
    wlist(list_words.ListIndex).jpn = txt_jpn.Text
    For i = 0 To opt_wtype.UBound
        If opt_wtype(i).Value Then wlist(list_words.ListIndex).wtype = i
    Next i
End Sub

Private Sub set_selection(obj As Object)
    obj.SetFocus
    obj.SelStart = 0
    obj.SelLength = Len(obj.Text)
End Sub
