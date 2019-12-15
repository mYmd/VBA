Attribute VB_Name = "Y_IO_utiliy"
'Y_IO_utility
'Copyright (c) 2016 mmYYmmdd
Option Explicit

'*********************************************************************************
'   IO�֘A���[�e�B���e�B
'*********************************************************************************
'   Function    sheet2m             Excel�V�[�g�̃Z���͈͂���z����擾
'   Sub         m2sheet             �z���Excel�V�[�g�̃Z���͈͂Ƀy�[�X�g
'   Function    getRangeMatrix      Excel�V�[�g�̃Z���͈͂���Range�I�u�W�F�N�g�̔z����擾
'   Function    getInterior         �I�u�W�F�N�g��Interior�v���p�e�B���擾
'   Function    push_back_l         vh_stdVec�ɒl��push_back����t�@���N�^�i���j
'   Function    push_back_r         vh_stdVec�ɒl��push_back����t�@���N�^�i�E�j
'   Function    textFile2m          �e�L�X�g�t�@�C���̔z��ǂݍ���
'   Function    mapTextFile         �e�L�X�g�t�@�C���̍s�P�ʂł̏���
'   Function    m2textFile          �z��̃e�L�X�g�t�@�C����������
'   Function    getURLText          URL�Ŏw�肳�ꂽ�e�L�X�g�̔z��ǂݍ���
'   Function    urlEncode           URL�G���R�[�h
'   Function    urlDecode           URL�f�R�[�h
'   Sub         m2Clip              �z��i2�����ȉ��j���N���b�v�{�[�h�ɓ]������
'   Function    HTMLDocFromText     HTML�e�L�X�g�����HTMLDocument�I�u�W�F�N�g
'   Function    getTagsFromHTML     HTML�e�L�X�g����̃^�O���o
'   (Function p_getTitleFromHTMLAnchor)
'   Function    wTable2m            Word�̃e�[�u������z����擾
'   Function    downloadFiles       URL�Ŏw�肵���t�@�C���̃_�E�����[�h
'   Function    encodeBase64        �t�@�C����Base64�G���R�[�h�ibinary -> text�j
'   Function    decodeBase64        �t�@�C����Base64�f�R�[�h�itext -> binary�j
'   Function    str2Base64          �������Base64�G���R�[�h
'*********************************************************************************

Public Declare PtrSafe Function textfile2array Lib "mapM.dll" _
                                    (ByRef fileName As Variant, _
                                ByVal codepage As Long, _
                            ByVal Head_N As Long, _
                        ByVal Head_Cut As Long, _
                    ByVal toAarry As Boolean) As Variant

Public Declare PtrSafe Function textfile_for_each Lib "mapM.dll" _
                                    (ByRef fun As Variant, _
                                ByRef fileName As Variant, _
                            ByVal codepage As Long, _
                        ByVal Head_N As Long, _
                    ByVal Head_Cut As Long) As Long

Public Declare PtrSafe Function array2textfile Lib "mapM.dll" _
                                        (ByRef m As Variant, _
                                    ByRef fileName As Variant, _
                                ByVal codepage As Long, _
                            ByVal append As Boolean, _
                        ByVal CrLf As Boolean) As Long
                                    
' Excel�V�[�g�̃Z���͈͂���z����擾�i�l�̂݁j
' vec = True�F1�����z��
' vec = Fale�F2�����z��i�f�t�H���g�j
' LBound = 0 �̔z��ƂȂ�
Public Function sheet2m(ByVal r As Object, Optional ByVal vec As Boolean = False) As Variant
    If StrConv(Application.Name, vbLowerCase) Like "*excel*" And TypeName(r) = "Range" Then
        If r.Cells.count = 1 Then
            sheet2m = makeM(1, 1, r.value)
        Else
            sheet2m = r.value
        End If
        If vec Then sheet2m = vector(sheet2m)
        changeLBound sheet2m, 0
    End If
End Function

' �z���Excel�V�[�g�̃Z���͈͂Ƀy�[�X�g�i����̃Z�����w��j
' vertical = True�F1�����z����c�Ƀy�[�X�g����
Public Sub m2sheet(ByRef matrix As Variant, _
                   ByVal r As Object, _
                   Optional ByVal vertical As Boolean = False)
    If StrConv(Application.Name, vbLowerCase) Like "*excel*" And TypeName(r) = "Range" And 0 < sizeof(matrix) Then
        Dim Lotus123Key As Boolean
        Lotus123Key = Application.TransitionNavigKeys
        If Lotus123Key Then Application.TransitionNavigKeys = False     '������
        Select Case Dimension(matrix)
        Case 0:
            r.value = matrix
        Case 1:
            If vertical Then
                r.Resize(sizeof(matrix), 1).value = transpose(matrix)
            Else
                r.Resize(1, sizeof(matrix)).value = matrix
            End If
        Case 2:
            r.Resize(rowSize(matrix), colSize(matrix)).value = matrix
        End Select
        If Lotus123Key Then Application.TransitionNavigKeys = True
    End If
End Sub

' Excel�V�[�g�̃Z���͈͂���Range�I�u�W�F�N�g�̔z����擾
Public Function getRangeMatrix(ByVal r As Object) As Variant
    If StrConv(Application.Name, vbLowerCase) Like "*excel*" And TypeName(r) = "Range" Then
        Dim i As Long, j As Long, ret As Variant
        With r
            ret = makeM(.rows.count, .Columns.count)
            For i = 0 To rowSize(ret) - 1 Step 1
                For j = 0 To colSize(ret) - 1 Step 1
                    Set ret(i, j) = .Cells(i + 1, j + 1)
                Next j
            Next i
        End With
    swapVariant getRangeMatrix, ret
    End If
End Function

' �I�u�W�F�N�g��Interior�v���p�e�B���擾
Public Function getInterior(ByRef Ob As Variant, ByRef dummuy As Variant) As Variant
    Set getInterior = Ob.Interior
End Function
    Public Function p_getInterior(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getInterior = make_funPointer(AddressOf getInterior, firstParam, secondParam)
    End Function

' vh_stdVec�ɒl��push_back����t�@���N�^�i���j
Public Function push_back_l(ByRef v As Variant, ByRef x As Variant) As Variant
    Set push_back_l = v.push_back(x)
End Function
    Function p_push_back_l(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_push_back_l = make_funPointer(AddressOf push_back_l, firstParam, secondParam)
    End Function

' vh_stdVec�ɒl��push_back����t�@���N�^�i�E�j
Public Function push_back_r(ByRef x As Variant, ByRef v As Variant) As Variant
    Set push_back_r = v.push_back(x)
End Function
    Function p_push_back_r(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_push_back_r = make_funPointer(AddressOf push_back_r, firstParam, secondParam)
    End Function

'---------------------------------------------------------------
        ' �����ϊ��p�R�[�h�y�[�W�iPrivate�I�Ȋ֐��j
        Function getCodePage(ByVal cp As String) As Long
            Select Case StrConv(cp, vbUpperCase + vbNarrow)
            Case "UTF-16", "UTF16", "UTF-16-LE", "UTF16-LE":
                                                getCodePage = 1200
            Case "UTF-8", "UTF8":               getCodePage = 65001
            Case "UTF-8N", "UTF8N":             getCodePage = -65001
            Case "UTF-7", "UTF7":               getCodePage = 65000
            Case "ANSI":                        getCodePage = 1252  'ANSI(1252)
            Case "SJIS", "S-JIS", "SHIFT-JIS":  getCodePage = 932
            Case "MAC":                         getCodePage = 2
            Case "OEM":                         getCodePage = 1
            Case "SYMBOL", "SYMBOL42":          getCodePage = 42
            Case Else
                If IsNumeric(cp) Then getCodePage = CLng(cp) Else getCodePage = 65001
            End Select
        End Function
'---------------------------------------------------------------

' �e�L�X�g�t�@�C���̔z��ǂݍ���
' codepage : �����������W���̖���
' head_n   : �����ǂݐ擪�s���w��
' head_cut : �擪�X�L�b�v�s���w��
' to_Array = False �̎��̓e�L�X�g�S�̂��ЂƂ̕�����ŕԂ�
Public Function textFile2m(ByVal fileName As String, _
                        ByVal codepage As String, _
                        Optional ByVal Head_N As Long = -1, _
                        Optional ByVal Head_Cut As Long = 0, _
                        Optional ByVal to_Array As Boolean = True) As Variant
    textFile2m = textfile2array(fileName, _
                            getCodePage(codepage), _
                        Head_N, _
                    Head_Cut, _
                to_Array)
End Function

' �e�L�X�g�t�@�C���̍s�P�ʂł̏���
' fn       : VBAHaskell�̊֐��i�L���Ȋ֐��łȂ���΍s�������Ԃ��j
' codepage : �����������W���̖���
' head_n   : �擪�s���w��
' head_cut : �擪�X�L�b�v�s���w��
' �߂�l   : �����s��
Public Function mapTextFile(ByRef fn As Variant, _
                            ByVal fileName As String, _
                            ByVal codepage As String, _
                            Optional ByVal Head_N As Long = -1, _
                            Optional ByVal Head_Cut As Long = 0) As Long
    mapTextFile = textfile_for_each(fn, _
                                fileName, _
                            getCodePage(codepage), _
                        Head_N, _
                    Head_Cut)
End Function

' �z��̃e�L�X�g�t�@�C����������
Public Function m2textFile(ByRef data As Variant, _
                           ByVal fileName As String, _
                           ByVal codepage As String, _
                           Optional ByVal append As Boolean = False, _
                           Optional ByVal CrLf As Boolean = True) As Long
    Dim codepage_ As Long
    codepage_ = getCodePage(codepage)
    If IsArray(data) Then
        m2textFile = array2textfile(data, fileName, codepage_, append, CrLf)
    Else
        m2textFile = array2textfile(Array(data), fileName, codepage_, append, CrLf)
    End If
End Function

' URL�Ŏw�肳�ꂽ�e�L�X�g�̔z��ǂݍ���
' head_n : �����ǂݐ擪�s���w��
Public Function getURLText(ByVal url As String, _
                           ByVal line_end As String, _
                           Optional ByVal Charset As String = "_autodetect_all", _
                           Optional ByVal Head_N As Long = -1, _
                           Optional ByVal userName As String = "", _
                           Optional ByVal passWord As String = "") As Variant
    
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo closeObjects
    With http
        If 0 < Len(userName) And 0 < Len(passWord) Then
            .Open "GET", url, False, userName, passWord
        Else
            .Open "GET", url, False
        End If
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Cache-Control", "no-cache"
        .Send
    End With
    Dim ado As Object:  Set ado = CreateObject("ADODB.Stream")
    Dim i As Long
    Dim lineS As String
    With ado
        .Open
        .Position = 0
        .Type = 1       'ADODB.Stream.adTypeBinary
        .Write http.responseBody
        .Position = 0
        .Type = 2       'ADODB.Stream.adTypeText
        .Charset = Charset
        If 0 < Head_N Then
            .LineSeparator = IIf(line_end = vbCr, Asc(vbCr), IIf(line_end = vbLf, Asc(vbLf), -1))
            For i = 1 To Head_N
                lineS = .ReadText(-2)   'adReadLine
                getURLText = getURLText & lineS & IIf(i = Head_N, "", line_end)
            Next i
        Else
            getURLText = .ReadText
        End If
    End With
closeObjects:
    If Not ado Is Nothing Then ado.Close
    Set ado = Nothing
    Set http = Nothing
    If 0 < Len(line_end) And VarType(getURLText) = vbString Then
        getURLText = Split(getURLText, line_end)
    End If
End Function

' URL�G���R�[�h�i�Q�l�����j
Public Function urlEncode(ByVal s As String) As String
    Dim ado As Object
    Dim tmp As Variant
    tmp = mapF(p_mid(s), zip(iota(1, Len(s)), repeat(1, Len(s))))
    Set ado = CreateObject("ADODB.Stream")
    ado.Charset = "UTF-8"
    tmp = mapF(p_urlEncode_1(, ado), tmp)
    Set ado = Nothing
    urlEncode = Join(tmp, "")
End Function

' URL�f�R�[�h�i�Q�l�����j
Public Function urlDecode(ByVal s As String) As String
    If s Like "*%??%??%??*" Then
        Dim begin As Long, theNext As Long
        begin = 1
        Dim ado As Object
        Set ado = CreateObject("ADODB.Stream")
        Do While begin <= Len(s) And Mid(s, begin) Like "*%??%??%??*"
            If Mid(s, begin, 9) Like "*%??%??%??*" Then
                urlDecode = urlDecode & urlDecode_imple(Mid(s, begin, 9), ado)
                begin = begin + 9
            Else
                theNext = InStr(begin + 1, s, "%")
                If 0 < theNext Then
                    urlDecode = urlDecode & Mid(s, begin, theNext - begin)
                    begin = theNext
                Else
                    urlDecode = urlDecode & Mid(s, begin)
                    begin = Len(s) + 1
                End If
            End If
        Loop
        Set ado = Nothing
        If begin < Len(s) Then
            urlDecode = urlDecode & Mid(s, begin)
        End If
    Else
        urlDecode = s
    End If
End Function

    Private Function isKanaKanji(ByVal s As String) As Boolean
        isKanaKanji = False
        If 0 < Len(s) Then
            If Left(s, 1) Like "[�-�]" Then
                isKanaKanji = True
            ElseIf Asc(Left(s, 1)) < 0 Then
                isKanaKanji = True
            End If
        End If
    End Function

    ' http://stabucky.com/wp/archives/4237
    Private Function urlEncode_1(ByRef s As Variant, ByRef adodbStream As Variant) As Variant
        Dim chars() As Byte
        If isKanaKanji(s) Then
            With adodbStream
                .Open
                .Type = 2       'ADODB.Stream.adTypeText
                .Position = 0
                .WriteText Left(s, 1)
                .Position = 0
                .Type = 1       'ADODB.Stream.adTypeBinary
                .Position = 3
                chars = .Read
                .Close
                urlEncode_1 = "%" & Hex(chars(0)) & "%" & Hex(chars(1)) & "%" & Hex(chars(2))
            End With
        Else
            urlEncode_1 = s
        End If
    End Function
    Private Function p_urlEncode_1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_urlEncode_1 = make_funPointer(AddressOf urlEncode_1, firstParam, secondParam)
    End Function


    Private Function urlDecode_imple(ByVal code As String, ByRef adodbStream As Object) As String
        Dim chars(0 To 2) As Byte
        chars(0) = CLng("&H" & Mid(code, 2, 2))
        chars(1) = CLng("&H" & Mid(code, 5, 2))
        chars(2) = CLng("&H" & Mid(code, 8, 2))
        With adodbStream
            .Open
            .Type = 1       'ADODB.Stream.adTypeBinary
            .Position = 0
            .Write chars
            .Position = 0
            .Type = 2       'ADODB.Stream.adTypeText
            .Charset = "UTF-8"
            urlDecode_imple = .ReadText
            .Close
        End With
    End Function

' �z��i2�����ȉ��j���N���b�v�{�[�h�ɓ]������
Public Sub m2Clip(ByRef data As Variant)
    Dim s As String
    Select Case Dimension(data)
    Case 0
        s = CStr_(data) & vbCrLf
    Case 1
        s = Join(mapF(p_CStr, data), vbTab) & vbCrLf
    Case 2
        Dim tmp As Variant
        tmp = zipC(mapF(p_CStr, data))
        tmp = mapF(p_join(, vbTab), tmp)
        s = Join(tmp, vbCrLf) & vbCrLf
    End Select
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .text = s
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

' HTML�e�L�X�g�����HTMLDocument�I�u�W�F�N�g
Public Function HTMLDocFromText(ByVal htmlText As String) As Object    'As HtmlDocument
    Set HTMLDocFromText = CreateObject("htmlfile") 'New MSHTML.HTMLDocument
    HTMLDocFromText.Write htmlText
End Function

' HTML�e�L�X�g����̃^�O���o
' �v���p�e�B�͈���prop�Ŏw��
' innerText, outerText, innerHTML, outerHTML, href �͕�����Ŏw��
' ����ȊO�� p_getTitleFromHTMLAnchor ���Q��
Public Function getTagsFromHTML(ByVal htmlText As String, _
                                ByVal tag As Variant, _
                                ByVal prop As Variant) As Variant
    Dim doc As Object   'As HTMLDocument
    Set doc = HTMLDocFromText(htmlText)
    Dim vec As vh_stdvec:   Set vec = New vh_stdvec
    Dim z As Variant
    If is_bindFun(prop) Then
        For Each z In doc.all.tags(tag)
            vec.push_back z
        Next z
        getTagsFromHTML = mapF(prop, vec.free)
    ElseIf TypeName(prop) = "String" Then
        Select Case StrConv(prop, vbLowerCase)
        Case "it"
            For Each z In doc.all.tags(tag)
                vec.push_back z.innerText
            Next z
        Case "ot"
            For Each z In doc.all.tags(tag)
                vec.push_back z.outerText
            Next z
        Case "ih"
            For Each z In doc.all.tags(tag)
                vec.push_back z.innerHTML
            Next z
        Case "oh"
            For Each z In doc.all.tags(tag)
                vec.push_back z.outerHTML
            Next z
        Case "href"
            For Each z In doc.all.tags(tag)
                vec.push_back z.href
            Next z
        End Select
        getTagsFromHTML = vec.free
    End If
    Set doc = Nothing
End Function

'HTMLAnchorElement ����'Title'�v���p�e�B���擾
    Private Function getTitleFromHTMLAnchor(ByRef anchor As Variant, _
                                            ByRef dummy As Variant) As Variant
        If TypeName(anchor) = "HTMLAnchorElement" Then
            getTitleFromHTMLAnchor = anchor.Title
        End If
    End Function
Public Function p_getTitleFromHTMLAnchor(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
    p_getTitleFromHTMLAnchor = make_funPointer(AddressOf getTitleFromHTMLAnchor, firstParam, secondParam)
End Function
    
' Word�̃e�[�u������z����擾
Public Function wTable2m(ByVal t As Object, Optional ByVal cutHeader As Boolean = False) As Variant
    If StrConv(Application.Name, vbLowerCase) Like "*word*" And TypeName(t) = "Table" Then
        Dim ret As Variant, tmp As String
        With t
            ret = makeM(.rows.count, .Columns.count)
            Dim i As Long, j As Long
            For i = 1 To .rows.count Step 1
                For j = 1 To .Columns.count Step 1
                    tmp = .Cell(i, j).Range.text
                    ret(i - 1, j - 1) = Left(tmp, Len(tmp) - 2)
                Next j
            Next i
        End With
    End If
    If cutHeader Then
        wTable2m = subM(ret, iota(1, UBound(ret, 1)))
    Else
        Call swapVariant(wTable2m, ret)
    End If
End Function

' URL�Ŏw�肵���t�@�C���̃_�E�����[�h
' fileNames�ȗ����̓����N���̒ʂ�i�Ō��'/'�̌�땔���j
' �߂�l�̓_�E�����[�h�ł����t�@�C���̐�
Public Function downloadFiles(ByRef URLs As Variant, _
                              ByVal pathName As String, _
                              Optional ByRef fileNames As Variant, _
                              Optional ByVal userName As String = "", _
                              Optional ByVal passWord As String = "") As Long
    downloadFiles = 0
    Dim fso As Object:      Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.folderExists(pathName) Then
        Set fso = Nothing
        Exit Function
    End If
    Dim URLm As Variant:    URLm = vector(URLs)
    Dim fNames As Variant
    If IsMissing(fileNames) Or IsEmpty(fileNames) Or IsNull(fileNames) Then
        fNames = mapF(p_getNth_b(, -1), mapF(p_split(, "/"), URLm))
    Else
        fNames = vector(fileNames)
    End If
    Dim i As Long
    For i = 0 To min_fun(UBound(URLm), UBound(fNames)) Step 1
        If downloadFile_(URLm(i), fso.BuildPath(pathName, fNames(i)), userName, passWord) Then
            downloadFiles = downloadFiles + 1
        End If
    Next i
    Set fso = Nothing
End Function

    Private Function downloadFile_(ByVal url As String, _
                                   ByVal fileName As String, _
                                   ByVal userName As String, _
                                   ByVal passWord As String) As Boolean
        downloadFile_ = False
        Dim oo As Object:   Set oo = CreateObject("MSXML2.XMLHTTP")
        If oo Is Nothing Then Exit Function
        On Error GoTo closeFun
        With oo
            If 0 < Len(userName) And 0 < Len(passWord) Then
                .Open "GET", url, False, userName, passWord
            Else
                .Open "GET", url, False
            End If
            .setRequestHeader "Pragma", "no-cache"
            .setRequestHeader "Cache-Control", "no-cache"
            .Send
        End With
        Select Case oo.Status       ' HTTP Result Code
        Case Is <= 200          ' ���N�G�X�g����
            With CreateObject("ADODB.Stream")
                .Type = 1 ' ADODB.Stream.adTypeBinary
                .Open
                .Write oo.responseBody
                .SaveToFile fileName, 2 ' adSaveCreateOverWrite
                .Close
            End With
            downloadFile_ = True
        End Select
closeFun:
    End Function

' �t�@�C����Base64�G���R�[�h�ibinary -> text�j
Function encodeBase64(ByVal fromFile As String, ByVal toFile As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If Not .FileExists(fromFile) Then Exit Function
        If Not .folderExists(.GetParentFolderName(toFile)) Then Exit Function
    End With
    Dim elm As Object
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").CreateElement("base64")
    With CreateObject("ADODB.Stream")
        .Type = 1   'adTypeBinary
        .Open
        .LoadFromFile fromFile
        elm.DataType = "bin.base64"
        elm.nodeTypedValue = .Read(-1)  'adReadAll
        .Close
    End With
    With CreateObject("ADODB.Stream")
        .Open
        .Position = 0
        .Type = 2    'ADODB.Stream.adTypeText
        .Charset = "UTF-8"
        .WriteText elm.text, 0          ' adWriteChar
        .SaveToFile toFile, 2   ' adSaveCreateOverWrite
        .Close
    End With
    encodeBase64 = True
End Function
 
' �t�@�C����Base64�f�R�[�h�itext -> binary�j
Function decodeBase64(ByVal fromFile As String, ByVal toFile As String) As Boolean
    With CreateObject("Scripting.FileSystemObject")
        If Not .FileExists(fromFile) Then Exit Function
        If Not .folderExists(.GetParentFolderName(toFile)) Then Exit Function
    End With
    Dim elm As Object
    Dim sss As String
    On Error Resume Next
    Set elm = CreateObject("MSXML2.DOMDocument").CreateElement("base64")
    With CreateObject("ADODB.Stream")
        .Open
        .Position = 0
        .Type = 2   'adTypeText
        .Charset = "UTF-8"
        .LoadFromFile fromFile
        elm.DataType = "bin.base64"
        elm.text = .ReadText  'adReadAll
        .Close
    End With
    With CreateObject("ADODB.Stream")
        .Type = 1   'adTypeBinary
        .Open
        .Write elm.nodeTypedValue
        .SaveToFile toFile, 2         'adSaveCreateOverWrite
        .Close
    End With
    decodeBase64 = True
End Function

' �������Base64�G���R�[�h
Function str2Base64(ByVal text As String, _
                        Optional ByVal charactor_set As String = "ASCII") As String
    Dim ado As Object:  Set ado = CreateObject("ADODB.Stream")
    With ado
        .Open
        .Position = 0
        .Charset = charactor_set
        .Type = 2   'adTypeTest
        .WriteText text
        .Position = 0
        .Type = 1   'adTypeBinary
    End With
    Dim dom As Object:  Set dom = CreateObject("MSXML2.DOMDocument").CreateElement("base64")
    With dom
        .DataType = "bin.base64"
        .nodeTypedValue = ado.Read
        str2Base64 = .text
    End With
    ado.Close
    Set ado = Nothing
    Set dom = Nothing
End Function
