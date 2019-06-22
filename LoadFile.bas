Attribute VB_Name = "LoadFile"
Option Explicit

Public Type Unit
    fpath As String
    uname As String
End Type
Public Type Word
    chn As String
    hira As String
    jpn As String
    wtype As Integer
End Type
Const INFOFILE As String = "\db.info"
Const DATAEXT As String = ".data"
Const SEPARATOR As String = "____"
Const DESCRIBESTR As String = ",[名] ,[代] ,[疑] ,[动1] ,[动2] ,[动3] ,[形1] ,[形2] ,[副] ,[连体] ,[连] ,[叹] ,[专] "
Public TYPE_DESCRIBE(13) As String
Public DATABASE As String
Public ulist() As Unit, wlist() As Word, data_file_name() As String

Public Function load_unit(Index As Integer) As Boolean
    If is_valid_unit(ulist(Index).fpath) Then
        Open ulist(Index).fpath For Input As #1
        Dim line As String, items
        Dim i As Integer
        i = 0
        ReDim wlist(0)
        Do While Not EOF(1)
            Input #1, line
            items = Split(line, SEPARATOR)
            If UBound(items) - LBound(items) = 3 Then
                ReDim Preserve wlist(i) As Word
                wlist(i).chn = items(0)
                wlist(i).hira = IIf(items(1) = "＊", "", items(1))
                wlist(i).jpn = items(2)
                wlist(i).wtype = Int(items(3))
                i = i + 1
            End If
        Loop
        Close #1
        load_unit = True
    Else
        MsgBox "Not a valid unit!" & vbCrLf & "Check if the file is missing.", _
            vbInformation, "File not found or damaged!"
        load_unit = False
    End If
End Function

Public Sub create_unit(unitname As String)
    
End Sub

Public Sub shuffle_words()
    If UBound(wlist) > 0 Then
        Dim temp As Integer, i As Integer
        Dim temp_word As Word
        For i = UBound(wlist) To 1 Step -1
            temp = Rnd * i
            temp_word = wlist(i)
            wlist(i) = wlist(temp)
            wlist(temp) = temp_word
        Next i
    End If
End Sub

Public Function load_database(filename As String) As Boolean
    load_database = is_valid_db(filename)
    If load_database Then
        DATABASE = filename
        Open filename & INFOFILE For Input As #1
        Dim line As String, items
        Dim i As Integer
        i = 0
        ReDim ulist(0)
        Do While Not EOF(1)
            Input #1, line
            items = Split(line, SEPARATOR)
            If UBound(items) - LBound(items) = 1 Then
                ReDim Preserve ulist(i)
                ulist(i).fpath = filename & "\" & items(0)
                ulist(i).uname = items(1)
                i = i + 1
            End If
        Loop
        Close #1
    End If
    items = Split(DESCRIBESTR, ",")
    For i = LBound(items) To UBound(items)
        TYPE_DESCRIBE(i) = items(i)
    Next i
End Function

Public Sub create_database(filepath As String)
    
End Sub

Public Function is_valid_unit(fname As String) As Boolean
    is_valid_unit = True
    If Dir(fname) = "" Then is_valid_unit = False
End Function

Public Function is_valid_db(filename As String) As Boolean
    MsgBox Dir(filename & INFOFILE)
    If Right(filename, 3) <> ".db" Then
        is_valid_db = False
    ElseIf Dir(filename & INFOFILE) = "" Then
        is_valid_db = False
    Else
        is_valid_db = True
    End If
End Function

Public Sub save_title_info(idx As Integer, fname As String, title As String)
    On Error GoTo RenameErr
    Name ulist(idx).fpath As DATABASE & "\" & fname & DATAEXT
    ulist(idx).fpath = DATABASE & "\" & fname & DATAEXT
    data_file_name(idx) = fname
    ulist(idx).uname = title
    save_unit_list
    Exit Sub
RenameErr:
    MsgBox "failed to rename file"
End Sub

Public Sub save_unit_list()
    Dim i As Integer
    On Error GoTo no_save_unit
    Open DATABASE & INFOFILE For Output As #1
    For i = 0 To UBound(ulist)
        Print #1, data_file_name(i) & DATAEXT & SEPARATOR & ulist(i).uname
    Next i
no_save_unit:
    Close #1
End Sub

Public Sub save_word_list(idx As Integer)
    On Error GoTo no_save_word
    Open ulist(idx).fpath For Output As #1
    Dim i As Integer
    For i = 0 To UBound(wlist)
        Print #1, wlist(i).chn & SEPARATOR & IIf(wlist(i).hira = "", "＊", wlist(i).hira) _
        & SEPARATOR & wlist(i).jpn & SEPARATOR & wlist(i).wtype
    Next i
no_save_word:
    Close #1
End Sub

Public Sub touch_file(fpath As String)
    If Dir(fpath) = "" Then
        Open fpath For Output As #1
        Close #1
    End If
End Sub

Public Sub swap_str(ByRef a As String, ByRef b As String)
    Dim t As String
    t = a
    a = b
    b = t
End Sub

Public Sub swap_unit(ByRef a As Unit, ByRef b As Unit)
    Dim t As Unit
    t = a
    a = b
    b = t
End Sub

Public Sub swap_word(ByRef a As Word, ByRef b As Word)
    Dim t As Word
    t = a
    a = b
    b = t
End Sub

