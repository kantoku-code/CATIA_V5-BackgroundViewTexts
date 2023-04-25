Language = "VBSCRIPT"
'*********************************
'BackgroundViewTexts.vbs"
'ver 0.01

'1-CATIA���N��
'2-����ٸد� �� Cat-Dll_Env-Path.txt���쐬
'3-BackgroundViewTexts.vbs
'  BackgroundViewTexts.CATScript
'  Cat-Dll_Env-Path.txt
'  �𓯈�̫��ނɒu��
'4-BackgroundViewTexts.vbs ��CATDrawinģ�ق�D&D
'5-CATDrawinģ�قƓ���̫��ނ�CSV̧�ق��o���オ��
'*********************************

'Option Explicit
'*** �ݒ� �ύX���Ȃ��ŉ����� ***
Const SCRIPT_NAME = "BackgroundViewTexts"
Const ENV_NAME = "Cat-Dll_Env-Path.txt"
Const LIST_FILE_NAME = "BackgroundViewTexts_IptLst.txt"
Const MACRO_NAME = "BackgroundViewTexts.CATScript"
Const ENV_KEYS = "CAT,DIRENV,ENV,CATTEMP"
'***

Call Main
wscript.Quit 0

'*********************************
Sub Main()
    'D&D
    Dim dDlist 'As Variant
    dDlist = get_drop_list(wscript.Arguments)
    If Not IsArray(dDlist) Then Exit Sub '���쐬��Exit

    '���߽�ގ擾
    Dim envDic 'As Object
    Set envDic = get_dll_env
    If envDic Is Nothing Then Exit Sub

    '�m�F
    Dim dDlistStr 'as String

    Dim msg 'As String
    msg = "�ȉ���̧�ق���e�L�X�g�f�[�^�𒊏o���܂��B��낵���ł����H" + vbNewLine + _
      dDList_to_string(dDlist)
    If MsgBox(msg, vbYesNo, SCRIPT_NAME) = vbNo Then Exit Sub

    '���߰�ؽč쐬
    Dim listPath 'As String
    listPath = Replace(envDic("CATTEMP"), Chr(34), "") + "\" + LIST_FILE_NAME
    If is_exists(listPath) Then
        msg = "�������̉\�����L��܂��B" + vbNewLine + _
          "�����I�Ɏ��s���܂����H"
        If MsgBox(msg, vbYesNo, SCRIPT_NAME) = vbNo Then Exit Sub
    End If
    Call write_file(listPath, Join(dDlist, vbNewLine))

    '�ޯ�Ӱ�ދN��
    Dim macroPath 'As String
    macroPath = get_current_path + "\" + MACRO_NAME

    Call execute_butch_mode(envDic, macroPath)

    '��ͻ�߰Ľ��ظĂŏ���
End Sub


' *** ButchMode ***
Private Sub execute_butch_mode(ByVal Dic, ByVal macroPath)
    Dim command 'As String
    command = Dic("CAT") & " -direnv " & _
        Dic("DIRENV") & " -env " & _
        Dic("ENV") & " -batch  -macro " & _
        Chr(34) & macroPath & Chr(34)

    Call CreateObject("Wscript.Shell").Exec(command)
End Sub


' *** Env ***
'�ޯ�Ӱ�ދN���p̧�َ擾
Private Function get_dll_env() 'As Object

    Dim envPath 'As String
    envPath = get_current_path + "\" + ENV_NAME

    If Not is_exists(envPath) Then
        Dim msg 'As String
        msg = "�ޯ�Ӱ�ދN���ɕK�v��̧�ق�����܂���!" + _
            vbNewLine + "(" + ENV_NAME + ")"
        MsgBox msg, vbOKOnly, SCRIPT_NAME
        Set get_dll_env = Nothing
        Exit Function
    End If

    Dim txts 'As Variant
    txts = read_file(envPath)
    If UBound(txts) < 3 Then Exit Function

    Dim Dic 'As Object
    Set Dic = CreateObject("Scripting.Dictionary")

    Dim i 'As Long
    Dim keyValue 'As Variant
    For i = 0 To UBound(txts)
        keyValue = get_key_value(txts(i))
        If Not UBound(keyValue) = 1 Then Exit Function
        Dic.Add keyValue(0), keyValue(1)
    Next

    If Not check_env(Dic) Then
        Set get_dll_env = Nothing
        Exit Function
    End If
    Set get_dll_env = Dic

End Function


'��������
Private Function check_env(ByVal Dic) 'As Boolean

    Dim i 'As Long
    Dim aryENV_KEYS 'As Variant
    aryENV_KEYS = Split(ENV_KEYS, ",")

    For i = 0 To UBound(aryENV_KEYS)
        If Not Dic.Exists(aryENV_KEYS(i)) Then
            Dim msg 'As String
            msg = "�ޯ�Ӱ�ދN���ɕK�v��̧�ٓ��̐ݒ肪����܂���!" + _
                vbNewLine + "(" + aryENV_KEYS(i) + ")"
            MsgBox msg, vbOKOnly, SCRIPT_NAME
            check_env = False
            Exit Function
        End If
    Next
    check_env = True

End Function


'�N���pKeyValue
'Return: 0-Key 1-Value
Private Function get_key_value(ByVal txt) 'As Variant

    Dim equal 'As Variant
    equal = Split(txt, "=")
    If Not UBound(equal) = 1 Then Exit Function

    Dim spece 'As Variant
    spece = Split(equal(0), " ")
    If Not UBound(spece) = 1 Then Exit Function

    Dim keyValue(1) 'As String
    keyValue(0) = spece(1)
    keyValue(1) = equal(1)
    get_key_value = keyValue

End Function


' *** D&D ***
'��ۯ�ߏ���
Private Function get_drop_list(ByVal Args) 'As Variant
    Dim argsCount 'As Long
    argsCount = Args.count

    If argsCount < 1 Then
        Call get_env_main
        Exit Function
    End If

    Dim drawList() 'As Variant
    ReDim drawList(argsCount)

    Dim fileCount 'As Long
    fileCount = -1

    Dim i 'As Long
    Dim path 'As Variant
    Dim argsPath 'As String

    'Continue��Goto�g�������������
    For i = 1 To argsCount
        argsPath = Args(i - 1)
        If is_exists(argsPath) Then
            path = split_path_name(argsPath)
            If is_drawFile(path(2)) Then
                fileCount = fileCount + 1
                drawList(fileCount) = join_path_name(path)
            End If
        End If
    Next

    If fileCount < 0 Then
        msg = "�ϊ��\��̧�ق�����܂���!"
        MsgBox msg, vbOKOnly, SCRIPT_NAME
        Exit Function
    End If

    ReDim Preserve drawList(fileCount)
    get_drop_list = drawList

End Function


'Iges�����@�g���q�̂� iif()�g������
Private Function is_drawFile(ByVal Ext) 'As Boolean

    is_drawFile = False
    If UCase(Ext) = "CATDRAWING" Then is_drawFile = True

End Function


'ؽĂ�̧��Ҳ�̂ݎ擾
Private Function dDList_to_string(ByVal dDlist) 'As Boolean

    Dim ts, toStr, i
    toStr = ""
    For i = 0 To UBound(dDlist)
        ts = split_path_name(dDlist(i))
        toStr = toStr + ts(1) + "." + ts(2) + vbNewLine
    Next
    dDList_to_string = toStr

End Function


' *** IO ***
'FileSystemObject
Private Function get_fso() 'As Object

    Set get_fso = CreateObject("Scripting.FileSystemObject")

End Function


'�߽/̧�ٖ�/�g���q ����
'Return: 0-Path 1-BaseName 2-Extension
Private Function split_path_name(ByVal fullPath) 'As Variant

    Dim path(2) 'As String
    With get_fso
        path(0) = .GetParentFolderName(fullPath)
        path(1) = .GetBaseName(fullPath)
        path(2) = .GetExtensionName(fullPath)
    End With
    split_path_name = path

End Function


'�߽/̧�ٖ�/�g���q �A��
Private Function join_path_name(ByVal path) 'As String

    If Not IsArray(path) Then Stop '���Ή�
    If Not UBound(path) = 2 Then Stop '���Ή�
    join_path_name = path(0) + "\" + path(1) + "." + path(2)

End Function


'̧�ق̗L��
Private Function is_exists(ByVal path) 'As Boolean

    is_exists = get_fso.FileExists(path)

End Function


'̧�ٓǂݍ���
Private Function read_file(ByVal path) 'As Variant

    With get_fso.GetFile(path).OpenAsTextStream
        read_file = Split(.ReadAll, vbNewLine)
        .Close
    End With

End Function


'̧�ُ�������
Private Sub write_file(ByVal path, ByVal txt)
    With get_fso.OpenTextFile(path, 2, True)
        .Write txt
        .Close
    End With
End Sub


'�������߽
Private Function get_current_path() 'As String

    get_current_path = get_fso.GetParentFolderName(wscript.ScriptFullName)

End Function


' *** ���擾 ***
Private Sub get_env_main()

    '���߽���擾����CATIA�̎擾
    Dim cat 'As Application
    Set cat = get_catia()
    If cat Is Nothing Then Exit Sub

    'catia�̎��ş���߽�擾
    Dim catPath ' As String
    catPath = cat.SystemService.Environ("CATDLLPath")

    '��̧���߽�擾
    Dim environmentPath ' As Variant
    environmentPath = split_path_name(cat.SystemService.Environ("CATEnvName"))

    'TEMP̫����߽�擾
    Dim tempPath ' As Variant
    tempPath = cat.SystemService.Environ("CATTemp")

    '�o�͕���
    Dim expTxt ' As String
    expTxt = "Set CAT=" + Chr(34) + catPath + "\CNEXT.exe" + Chr(34) + vbNewLine + _
             "Set DIRENV=" + Chr(34) + environmentPath(0) + Chr(34) + vbNewLine + _
             "Set ENV=" + Chr(34) + environmentPath(1) + Chr(34) + vbNewLine + _
             "Set CATTEMP=" + Chr(34) + tempPath + Chr(34)

    '�ۑ�
    Dim expPath 'As String
    expPath = get_current_path + "\" + ENV_NAME
    If is_exists(expPath) Then
        Dim msg 'As String
        msg = "�u" + ENV_NAME + "�v�����݂��܂��B�㏑�����܂���?(������-��ݾ�)"
        If MsgBox(msg, vbYesNo, SCRIPT_NAME) = vbNo Then Exit Sub
    End If
    Call write_file(expPath, expTxt)

    '�I��
    MsgBox expPath + vbNewLine + "���쐬���܂���", vbOKOnly, SCRIPT_NAME

End Sub


'�N������catia�̎擾
Private Function get_catia() 'As Application

    On Error Resume Next
        Set get_catia = GetObject(, "CATIA.Application")
        If get_catia Is Nothing Then
            MsgBox "CATIA V5 ���N�����Ă�������", vbOKOnly, SCRIPT_NAME
            Err.Clear
            wscript.Quit 0
        End If
    On Error GoTo 0

End Function