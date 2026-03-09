Attribute VB_Name = "Module_Markdown"
Option Compare Database
Option Explicit
 
'
' ǰ������:
'   1. ���� JsonConverter ģ�� (VBA-JSON by Tim Hall)
'   2. ���� -> ���� -> ��ѡ "Microsoft Scripting Runtime"
'   3. �޸� API_KEY Ϊ��� DeepSeek Key
'
' ���ٿ�ʼ:
'   �� VBA ��������ִ��:
'       CreateAIForm          ' �Զ����� AI �ʴ���
'   Ȼ���� Access �д򿪴��� frmAI, ��������, ��� [����]
'
'   ��������:
'       ShowMarkdown "# ����" & vbCrLf & "**����**"
'       MarkdownDemo
'====================================================

' ---------- Win32 Sleep ----------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------- ���� ----------
Private Const API_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Private Const API_URL   As String = "https://api.deepseek.com/chat/completions"
Private Const API_MODEL As String = "deepseek-chat"

Private Const AI_FORM   As String = "frmAI"
Private Const MD_FORM   As String = "frmMarkdownViewer"
Private Const TXT_MD    As String = "txtMarkdown"


'############################################################
'#                                                          #
'#   ��һ����: Markdown -> ���ı� HTML                       #
'#                                                          #
'############################################################

'====================================================
' ����: Markdown -> Access ���ı� HTML
' Access ֧��: <b> <i> <u> <p> <br> <font> <ul> <ol> <li>
'====================================================
Public Function MarkdownToRichText(ByVal sMd As String) As String
    Dim vLines As Variant
    Dim out As String
    Dim i As Long
    Dim ln As String
    Dim inCode As Boolean
    Dim inUL As Boolean
    Dim inOL As Boolean

    sMd = Replace(sMd, vbCrLf, vbLf)
    sMd = Replace(sMd, vbCr, vbLf)
    vLines = Split(sMd, vbLf)

    out = ""

    For i = 0 To UBound(vLines)
        ln = CStr(vLines(i))

        ' ---- ����� ``` ----
        If Left$(Trim$(ln), 3) = "```" Then
            If Not inCode Then
                If inUL Then: out = out & "</ul>": inUL = False
                If inOL Then: out = out & "</ol>": inOL = False
                inCode = True
            Else
                inCode = False
            End If
            GoTo NxtLine
        End If
        If inCode Then
            out = out & "<p><font face=""Consolas"" size=""2"" color=""#333333"">" & EscHtml(ln) & "</font></p>"
            GoTo NxtLine
        End If

        ' ---- ���� ----
        If Len(Trim$(ln)) = 0 Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            GoTo NxtLine
        End If

        ' ---- ˮƽ�� ----
        If IsHRule(ln) Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            out = out & "<p><font color=""#cccccc"">" & String$(40, "-") & "</font></p>"
            GoTo NxtLine
        End If

        ' ---- ���� # ~ ###### ----
        Dim hLv As Long
        hLv = HeadingLevel(ln)
        If hLv > 0 Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            Dim hTxt As String
            hTxt = Trim$(Mid$(ln, hLv + 1))
            Do While Len(hTxt) > 0 And Right$(hTxt, 1) = "#"
                hTxt = RTrim$(Left$(hTxt, Len(hTxt) - 1))
            Loop
            Dim fs As Long
            Select Case hLv
                Case 1: fs = 7
                Case 2: fs = 6
                Case 3: fs = 5
                Case 4: fs = 4
                Case Else: fs = 3
            End Select
            out = out & "<p><font size=""" & fs & """><b>" & FmtInline(hTxt) & "</b></font></p>"
            GoTo NxtLine
        End If

        ' ---- ���� > ----
        If Left$(LTrim$(ln), 1) = ">" Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            Dim qTxt As String
            qTxt = LTrim$(ln)
            If Left$(qTxt, 2) = "> " Then
                qTxt = Mid$(qTxt, 3)
            ElseIf Left$(qTxt, 1) = ">" Then
                qTxt = Mid$(qTxt, 2)
            End If
            out = out & "<p><font color=""#57606a"">| " & FmtInline(qTxt) & "</font></p>"
            GoTo NxtLine
        End If

        ' ---- �����б� ----
        If IsULItem(ln) Then
            If inOL Then: out = out & "</ol>": inOL = False
            If Not inUL Then: out = out & "<ul>": inUL = True
            out = out & "<li>" & FmtInline(ULItemText(ln)) & "</li>"
            GoTo NxtLine
        Else
            If inUL Then: out = out & "</ul>": inUL = False
        End If

        ' ---- �����б� ----
        If IsOLItem(ln) Then
            If inUL Then: out = out & "</ul>": inUL = False
            If Not inOL Then: out = out & "<ol>": inOL = True
            out = out & "<li>" & FmtInline(OLItemText(ln)) & "</li>"
            GoTo NxtLine
        Else
            If inOL Then: out = out & "</ol>": inOL = False
        End If

        ' ---- ���� ----
        If IsTblRow(ln) Then
            If Not IsTblSep(ln) Then
                out = out & "<p><font face=""Consolas"" size=""2"">" & EscHtml(ln) & "</font></p>"
            End If
            GoTo NxtLine
        End If

        ' ---- ��ͨ���� ----
        out = out & "<p>" & FmtInline(ln) & "</p>"

NxtLine:
    Next i

    If inUL Then out = out & "</ul>"
    If inOL Then out = out & "</ol>"

    MarkdownToRichText = out
End Function

'====================================================
' ���ڸ�ʽ (����/б��/����/����/ͼƬ/ɾ����)
'====================================================
Private Function FmtInline(ByVal s As String) As String
    Dim re As Object

    s = EscHtml(s)

    ' `code`
    Set re = MakeRE("`([^`]+)`")
    s = re.Replace(s, "<font face=""Consolas"" color=""#c7254e"">$1</font>")

    ' ![alt](url) - ͼƬ (����������֮ǰ)
    Set re = MakeRE("!\[([^\]]*)\]\(([^)]+)\)")
    s = re.Replace(s, "<font color=""#999999"">[img: $1]</font>")

    ' [text](url) - ����
    Set re = MakeRE("\[([^\]]+)\]\(([^)]+)\)")
    s = re.Replace(s, "<font color=""#0366d6""><u>$1</u></font>")

    ' ***text*** / ___text___
    Set re = MakeRE("\*{3}(.+?)\*{3}")
    s = re.Replace(s, "<b><i>$1</i></b>")
    Set re = MakeRE("_{3}(.+?)_{3}")
    s = re.Replace(s, "<b><i>$1</i></b>")

    ' **text** / __text__
    Set re = MakeRE("\*{2}(.+?)\*{2}")
    s = re.Replace(s, "<b>$1</b>")
    Set re = MakeRE("_{2}(.+?)_{2}")
    s = re.Replace(s, "<b>$1</b>")

    ' *text* / _text_
    Set re = MakeRE("\*(.+?)\*")
    s = re.Replace(s, "<i>$1</i>")
    Set re = MakeRE("\b_(.+?)_\b")
    s = re.Replace(s, "<i>$1</i>")

    ' ~~text~~
    Set re = MakeRE("~~(.+?)~~")
    s = re.Replace(s, "<font color=""#999999"">$1</font>")

    FmtInline = s
End Function


'############################################################
'#                                                          #
'#   �ڶ�����: DeepSeek API ����                             #
'#   ����A: curl �ӽ�������ʽ (Windows 10 1803+)            #
'#   ����B: ͬ������ + ���ֻ�Ч�� (����)                     #
'#                                                          #
'############################################################

'====================================================
' ��ť���
'====================================================
Public Sub Askai()
    Dim frm As Form
    Set frm = Screen.ActiveForm

    If Len(Trim$(Nz(frm!txtQ, ""))) = 0 Then
        MsgBox "���������⡣", vbInformation
        Exit Sub
    End If

    Dim sQuestion As String
    sQuestion = CStr(frm!txtQ)

    ' curl.exe �� Windows 10 1803 ��ʼ����
    If Dir(Environ$("SystemRoot") & "\System32\curl.exe") <> "" Then
        StreamWithCurl frm, sQuestion
    Else
        SyncWithTypewriter frm, sQuestion
    End If
End Sub

'====================================================
' ����A: ����ʽ �� curl �ӽ����� SSE ����
'
' ԭ��:
'   1. ��������д����ʱ JSON �ļ�
'   2. �� Shell ���� curl, �� SSE ��ʽ����, �������ʱ�ļ�
'   3. VBA ÿ 80ms ��ѯ��ʱ�ļ�, ��ȡ��������
'   4. ���� SSE data ��, ��ȡ delta.content
'   5. ʵʱ�����ı��� (�����ı߽��ձ���ʾ)
'   6. �յ� [DONE] ��תΪ Markdown ���ı�
'====================================================
Private Sub StreamWithCurl(frm As Form, ByVal sQuestion As String)
    On Error GoTo ErrHandler

    ' --- ׼����ʱ�ļ� ---
    Dim sTS As String
    Randomize
    sTS = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000))
    Dim sTmpBody As String
    Dim sTmpResp As String
    Dim sTmpErr As String
    Dim sTmpDone As String
    sTmpBody = Environ$("TEMP") & "\ds_body_" & sTS & ".json"
    sTmpResp = Environ$("TEMP") & "\ds_resp_" & sTS & ".txt"
    sTmpErr = Environ$("TEMP") & "\ds_err_" & sTS & ".txt"
    sTmpDone = Environ$("TEMP") & "\ds_done_" & sTS & ".flag"

    ' ���������� (stream=true)
    Dim sBody As String
    sBody = BuildDeepSeekRequestBody(sQuestion, True)

    ' д���������ļ� (UTF-8 �� BOM)
    WriteUTF8NoBom sTmpBody, sBody

    ' ɾ������Ӧ�ļ�
    On Error Resume Next
    Kill sTmpResp
        Kill sTmpErr
        Kill sTmpDone
    Err.Clear
    On Error GoTo ErrHandler

    ' --- ���� curl ---
    Dim sCurl As String
        sCurl = """" & Environ$("SystemRoot") & "\System32\curl.exe"" " & _
                "--http1.1 -sS -N --no-buffer " & _
                "-X POST """ & API_URL & """ " & _
                "-H ""Content-Type: application/json; charset=utf-8"" " & _
                "-H ""Authorization: Bearer " & API_KEY & """ " & _
                "-H ""Accept: text/event-stream"" " & _
                "--data-binary @""" & sTmpBody & """"

        Dim sCmd As String
        sCmd = "cmd /c (" & sCurl & " 1>""" & sTmpResp & """ 2>""" & sTmpErr & """) & echo done>""" & sTmpDone & """"
        Shell sCmd, vbHide

    ' --- UI ��ʼ�� ---
    DoCmd.Hourglass True
    frm!lblMsg.Caption = "AI ����˼��..."
    frm!txtAnswer.TextFormat = acTextFormatPlain
    frm!txtAnswer.Value = ""
    frm.Repaint

    ' --- ��ѯ��Ӧ�ļ� ---
    Dim sFullText As String     ' �ۻ��������ش�
    Dim lLastRawLen As Long     ' �ϴζ�����ԭʼ�ı�����
    Dim sngStart As Single      ' ��ʼʱ��
    Dim sngLastUI As Single     ' �ϴ� UI ˢ��ʱ��
    Dim bDone As Boolean
    Dim bFirstToken As Boolean
    Dim bProcDone As Boolean
    Dim sAll As String
    Dim sErr As String
    Dim sCursor As String

    sFullText = ""
    lLastRawLen = 0
    sngStart = Timer
    sngLastUI = Timer
    bDone = False
    bFirstToken = False
    bProcDone = False
    sCursor = ChrW$(&H258C)    ' ��

    Do
        DoEvents
        Sleep 80                ' 80ms һ��

        ' ��ȡ��ʱ�ļ� (UTF-8)
        sAll = ReadFileAsUTF8(sTmpResp)

        ' ��������
        If Len(sAll) > lLastRawLen Then
            lLastRawLen = Len(sAll)

            ' ����Ƿ����
            If InStr(sAll, "[DONE]") > 0 Then bDone = True

            ' ���½���ȫ�� SSE ���� (�򵥿ɿ�, ���½ض�)
            Dim sNewFull As String
            sNewFull = ParseSSEChunk(sAll)

            If Len(sNewFull) > Len(sFullText) Then
                sFullText = sNewFull

                ' �״��յ�����
                If Not bFirstToken Then
                    bFirstToken = True
                    DoCmd.Hourglass False
                    frm!lblMsg.Caption = "�������..."
                End If

                ' ������ʾ
                frm!txtAnswer.Value = sFullText & sCursor
                frm.Repaint
                sngLastUI = Timer
            End If
        End If

        bProcDone = (Dir$(sTmpDone) <> "")
        If bDone Then Exit Do
        If bProcDone Then Exit Do

        ' ��ʱ 180 ��
        Dim sngElapsed As Single
        sngElapsed = Timer - sngStart
        If sngElapsed < 0 Then sngElapsed = sngElapsed + 86400  ' ����ҹ
        If sngElapsed > 180 Then
            frm!lblMsg.Caption = "����ʱ��"
            sErr = ReadFileAsUTF8(sTmpErr)
            If Len(sErr) > 0 Then
                MsgBox "����ʱ��curl ���:" & vbCrLf & Left$(sErr, 1000), vbExclamation
            Else
                MsgBox "����ʱ (180��)��", vbExclamation
            End If
            Exit Do
        End If
    Loop

    ' --- ������ʾ: Markdown ���ı� ---
    DoCmd.Hourglass False
    If Len(sFullText) > 0 Then
        frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
        frm!txtAnswer.Value = MarkdownToRichText(sFullText)
        frm!lblMsg.Caption = "�ش���ɡ� (�� " & Len(sFullText) & " �ַ�)"
    Else
        ' �����Ǵ�����Ӧ
        sAll = ReadFileAsUTF8(sTmpResp)
        sErr = ReadFileAsUTF8(sTmpErr)
        If Len(sErr) > 0 Then
            frm!txtAnswer.Value = "curl ����:" & vbCrLf & Left$(sErr, 1500)
            frm!lblMsg.Caption = "curl ִ��ʧ�ܡ�"
        ElseIf Len(sAll) > 0 Then
            frm!txtAnswer.Value = "����ʧ��:" & vbCrLf & Left$(sAll, 1000)
            frm!lblMsg.Caption = "��ɣ����������ݲ�����Ч SSE��"
        Else
            frm!txtAnswer.Value = "(δ�յ��ش�)"
            frm!lblMsg.Caption = "curl �ѽ�������δ�յ����ݡ�"
        End If
    End If
    frm.Repaint

    ' ������ʱ�ļ�
    On Error Resume Next
    Kill sTmpBody
    Kill sTmpResp
    Kill sTmpErr
    Kill sTmpDone
    On Error GoTo 0
    Exit Sub

ErrHandler:
    DoCmd.Hourglass False
    On Error Resume Next
    Kill sTmpBody
    Kill sTmpResp
    Kill sTmpErr
    Kill sTmpDone
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    On Error GoTo 0
    MsgBox "StreamWithCurl Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

'====================================================
' ����B ����: ͬ������ + ���ֻ�Ч��
' (curl ������ʱ�Զ�ʹ��)
'====================================================
Private Sub SyncWithTypewriter(frm As Form, ByVal sQuestion As String)
    On Error GoTo ErrHandler

    Dim sBody As String
    sBody = BuildDeepSeekRequestBody(sQuestion, False)

    DoCmd.Hourglass True
    frm!lblMsg.Caption = "AI ����˼��..."
    frm!txtAnswer.TextFormat = acTextFormatPlain
    frm!txtAnswer.Value = ""
    frm.Repaint

    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts 5000, 10000, 30000, 180000
    xmlHttp.Open "POST", API_URL, False
    xmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    xmlHttp.setRequestHeader "Authorization", "Bearer " & API_KEY
    xmlHttp.send sBody

    If xmlHttp.Status <> 200 Then
        frm!lblMsg.Caption = "����ʧ��: HTTP " & xmlHttp.Status
        MsgBox "HTTP " & xmlHttp.Status & vbCrLf & _
               Left$(xmlHttp.responseText, 500), vbExclamation
        GoTo ExitHere
    End If

    Dim oJson As Object
    Set oJson = JsonConverter.ParseJson(xmlHttp.responseText)
    Dim sAnswer As String
    sAnswer = oJson("choices")(1)("message")("content")
    Set xmlHttp = Nothing
    DoCmd.Hourglass False

    If Len(sAnswer) = 0 Then
        MsgBox "API ��������Ϊ�ա�", vbExclamation
        GoTo ExitHere
    End If

    frm!lblMsg.Caption = "�������..."
    TypewriterShow frm, sAnswer

    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = MarkdownToRichText(sAnswer)
    frm!lblMsg.Caption = "�ش���ɡ� (�� " & Len(sAnswer) & " �ַ�)"

ExitHere:
    DoCmd.Hourglass False
    Set xmlHttp = Nothing
    Exit Sub

ErrHandler:
    DoCmd.Hourglass False
    On Error Resume Next
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume ExitHere
End Sub

'====================================================
' ���ֻ�Ч�� (����B ʹ��, �ٶ�����Ӧ)
'====================================================
Private Sub TypewriterShow(frm As Form, ByVal sText As String)
    Dim lTotal As Long
    Dim lPos As Long
    Dim lStep As Long
    Dim lDelay As Long
    Dim sCursor As String

    sCursor = ChrW$(&H258C)
    lTotal = Len(sText)
    If lTotal = 0 Then Exit Sub

    If lTotal < 500 Then
        lStep = 2: lDelay = 25
    ElseIf lTotal < 1500 Then
        lStep = 4: lDelay = 20
    ElseIf lTotal < 3000 Then
        lStep = 8: lDelay = 15
    Else
        lStep = 15: lDelay = 10
    End If

    frm!txtAnswer.TextFormat = acTextFormatPlain

    For lPos = lStep To lTotal Step lStep
        frm!txtAnswer.Value = Left$(sText, lPos) & sCursor
        frm.Repaint
        DoEvents
        Sleep lDelay
    Next lPos

    frm!txtAnswer.Value = sText
    frm.Repaint
End Sub

'====================================================
' SSE ����: ��ȡ���� data �е� delta.content
'====================================================
Private Function ParseSSEChunk(ByVal sChunk As String) As String
    Dim vLines As Variant
    Dim i As Long
    Dim sLine As String
    Dim sJsonStr As String
    Dim sResult As String

    sChunk = Replace(sChunk, vbCrLf, vbLf)
    sChunk = Replace(sChunk, vbCr, vbLf)
    vLines = Split(sChunk, vbLf)

    sResult = ""
    For i = 0 To UBound(vLines)
        sLine = CStr(vLines(i))
        If Left$(sLine, 6) = "data: " Then
            sJsonStr = Mid$(sLine, 7)
            If sJsonStr <> "[DONE]" And Len(Trim$(sJsonStr)) > 0 Then
                sResult = sResult & ExtractDelta(sJsonStr)
            End If
        End If
    Next i

    ParseSSEChunk = sResult
End Function

'====================================================
' �� JsonConverter �������� SSE JSON
'====================================================
Private Function ExtractDelta(ByVal sJson As String) As String
    On Error Resume Next

    Dim oJson As Object
    Set oJson = JsonConverter.ParseJson(sJson)
    If Err.Number <> 0 Then
        Err.Clear
        ExtractDelta = ""
        Exit Function
    End If

    Dim sDelta As String
    sDelta = oJson("choices")(1)("delta")("content")
    If Err.Number <> 0 Then
        Err.Clear
        ExtractDelta = ""
        Exit Function
    End If

    ExtractDelta = sDelta
End Function

'====================================================
' ͳһ���� DeepSeek ������
' ʹ�� JsonConverter ���л�, �����ֹ�ƴ JSON ����
'====================================================
Private Function BuildDeepSeekRequestBody(ByVal sQuestion As String, _
                                          Optional ByVal bStream As Boolean = False) As String
    Dim oRoot As Object
    Dim oMsg As Object
    Dim colMessages As Collection

    Set oRoot = CreateObject("Scripting.Dictionary")
    Set oMsg = CreateObject("Scripting.Dictionary")
    Set colMessages = New Collection

    oMsg.Add "role", "user"
    oMsg.Add "content", sQuestion
    colMessages.Add oMsg

    oRoot.Add "model", API_MODEL
    oRoot.Add "messages", colMessages
    oRoot.Add "temperature", 0.7
    oRoot.Add "max_tokens", 8192
    If bStream Then oRoot.Add "stream", True

    BuildDeepSeekRequestBody = JsonConverter.ConvertToJson(oRoot)
End Function

'====================================================
' UTF-8 �ļ�д�� (�� BOM, curl ��Ҫ)
'====================================================
Private Sub WriteUTF8NoBom(ByVal sPath As String, ByVal sText As String)
    ' ���� ADODB.Stream д UTF-8 (��� BOM)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText sText
    stm.SaveToFile sPath, 2
    stm.Close
    Set stm = Nothing

    ' ���¶�ȡ������, ȥ�� 3 �ֽ� BOM (EF BB BF)
    Dim f As Integer
    Dim bAll() As Byte
    Dim lLen As Long

    f = FreeFile
    Open sPath For Binary Access Read As #f
    lLen = LOF(f)
    If lLen <= 3 Then
        Close #f
        Exit Sub
    End If
    ReDim bAll(lLen - 1)
    Get #f, 1, bAll
    Close #f

    ' ��� BOM
    If bAll(0) = &HEF And bAll(1) = &HBB And bAll(2) = &HBF Then
        ' ȥ��ǰ 3 �ֽ���д
        Dim bNoBom() As Byte
        ReDim bNoBom(lLen - 4)
        Dim j As Long
        For j = 0 To lLen - 4
            bNoBom(j) = bAll(j + 3)
        Next j
        ' ������ɾ�����ļ�, ���� Open For Binary ���ض�, �����β���ֽ�
        Kill sPath
        f = FreeFile
        Open sPath For Binary Access Write As #f
        Put #f, 1, bNoBom
        Close #f
    End If
End Sub

'====================================================
' ��ȡ��ʱ�ļ� (UTF-8 -> VBA �ַ���)
' �ļ��������� curl д��, ʧ��ʱ���ؿմ�
'====================================================
Private Function ReadFileAsUTF8(ByVal sPath As String) As String
    On Error Resume Next

    ' ����ļ��Ƿ����
    If Dir(sPath) = "" Then
        ReadFileAsUTF8 = ""
        Exit Function
    End If

    ' ��ȡԭʼ�ֽ�
    Dim f As Integer
    Dim lLen As Long
    Dim bArr() As Byte

    f = FreeFile
    Open sPath For Binary Access Read As #f
    If Err.Number <> 0 Then
        Err.Clear
        ReadFileAsUTF8 = ""
        Exit Function
    End If

    lLen = LOF(f)
    If lLen = 0 Then
        Close #f
        ReadFileAsUTF8 = ""
        Exit Function
    End If

    ReDim bArr(lLen - 1)
    Get #f, 1, bArr
    Close #f

    ' �� ADODB.Stream �� UTF-8 �ֽ�תΪ VBA �ַ���
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1    ' adTypeBinary
    stm.Open
    stm.Write bArr
    stm.Position = 0
    stm.Type = 2    ' adTypeText
    stm.Charset = "UTF-8"
    ReadFileAsUTF8 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing

    If Err.Number <> 0 Then
        Err.Clear
        ReadFileAsUTF8 = ""
    End If
End Function


'############################################################
'#                                                          #
'#   ��������: �����Զ�����                                  #
'#                                                          #
'############################################################

'====================================================
' ���� AI �ʴ��� frmAI
' ����: txtQ, txtAnswer(���ı�), lblMsg, btnAsk
'====================================================
Public Sub CreateAIForm()
    On Error GoTo Err_Create

    Dim frm As Form
    Dim ctl As Control
    Dim sTmp As String

    If FormExists(AI_FORM) Then
        DoCmd.Close acForm, AI_FORM, acSaveNo
        DoCmd.DeleteObject acForm, AI_FORM
    End If

    Set frm = CreateForm
    With frm
        .Caption = "AI �ʴ� (DeepSeek)"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .Width = 12000
        .Section(acDetail).Height = 10000
        .Section(acDetail).BackColor = RGB(248, 249, 250)
    End With

    ' --- ��ǩ: ���� ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 150, 1200, 350)
    ctl.Caption = "����:"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.FontBold = True

    ' --- txtQ: ��������� ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 200, 520, 9200, 600)
    ctl.Name = "txtQ"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.ScrollBars = 2
    ctl.EnterKeyBehavior = True

    ' --- btnAsk: ���ʰ�ť ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 9600, 520, 2000, 600)
    ctl.Name = "btnAsk"
    ctl.Caption = "��  ��"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.OnClick = "=btnAsk_Click()"

    ' --- lblMsg: ״̬��ǩ ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 1250, 11400, 350)
    ctl.Name = "lblMsg"
    ctl.Caption = "����������� [����]"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = RGB(108, 117, 125)

    ' --- txtAnswer: �ش���� (���ı�) ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 200, 1700, 11400, 8100)
    ctl.Name = "txtAnswer"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.ScrollBars = 2
    ctl.BackColor = RGB(255, 255, 255)
    ctl.BorderStyle = 1
    ctl.Locked = True
    ctl.TabStop = False
    ctl.EnterKeyBehavior = True

    ' ���洰��
    sTmp = frm.Name
    DoCmd.Close acForm, sTmp, acSaveYes
    Set frm = Nothing

    ' ���´������ͼ���� TextFormat (���뱣��������)
    DoCmd.OpenForm sTmp, acDesign
    Forms(sTmp).Controls("txtAnswer").TextFormat = acTextFormatHTMLRichText
    DoCmd.Close acForm, sTmp, acSaveYes

    ' ������
    If sTmp <> AI_FORM Then
        DoCmd.Rename AI_FORM, acForm, sTmp
    End If

    MsgBox "���� [" & AI_FORM & "] �����ɹ�!" & vbCrLf & vbCrLf & _
           "�򿪴��弴��ʹ�� AI �ʴ�", vbInformation
    Exit Sub

Err_Create:
    MsgBox "CreateAIForm: " & Err.Description, vbExclamation
End Sub

'====================================================
' ������ Markdown �鿴���� frmMarkdownViewer
'====================================================
Public Sub CreateMarkdownForm()
    On Error GoTo Err_Create

    Dim frm As Form
    Dim ctl As Control
    Dim sTmp As String

    If FormExists(MD_FORM) Then
        DoCmd.Close acForm, MD_FORM, acSaveNo
        DoCmd.DeleteObject acForm, MD_FORM
    End If

    Set frm = CreateForm
    With frm
        .Caption = "Markdown"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .Width = 11000
        .Section(acDetail).Height = 9000
        .Section(acDetail).BackColor = RGB(255, 255, 255)
    End With

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 30, 30, 10940, 8940)
    ctl.Name = TXT_MD
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.ScrollBars = 2
    ctl.BackColor = RGB(255, 255, 255)
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0
    ctl.Locked = True
    ctl.TabStop = False
    ctl.EnterKeyBehavior = True

    sTmp = frm.Name
    DoCmd.Close acForm, sTmp, acSaveYes
    Set frm = Nothing

    DoCmd.OpenForm sTmp, acDesign
    Forms(sTmp).Controls(TXT_MD).TextFormat = acTextFormatHTMLRichText
    DoCmd.Close acForm, sTmp, acSaveYes

    If sTmp <> MD_FORM Then
        DoCmd.Rename MD_FORM, acForm, sTmp
    End If

    Debug.Print "[" & MD_FORM & "] OK"
    Exit Sub

Err_Create:
    MsgBox "CreateMarkdownForm: " & Err.Description, vbExclamation
End Sub


'############################################################
'#                                                          #
'#   ���Ĳ���: ��ʾ/���ߺ���                                 #
'#                                                          #
'############################################################

'====================================================
' ������ʾ Markdown
'====================================================
Public Sub ShowMarkdown(ByVal sMd As String, Optional ByVal sTitle As String = "Markdown")
    On Error GoTo Err_Show

    If Not FormExists(MD_FORM) Then
        CreateMarkdownForm
    End If

    DoCmd.OpenForm MD_FORM, acNormal
    With Forms(MD_FORM)
        .Caption = sTitle
        .Controls(TXT_MD).Value = MarkdownToRichText(sMd)
    End With
    Exit Sub

Err_Show:
    MsgBox "ShowMarkdown: " & Err.Description, vbExclamation
End Sub

'====================================================
' д�����⸻�ı��ı���
'====================================================
Public Sub SetTextBoxMarkdown(txt As TextBox, ByVal sMd As String)
    txt.TextFormat = acTextFormatHTMLRichText
    txt.Value = MarkdownToRichText(sMd)
End Sub

'====================================================
' HTML ת��
'====================================================
Private Function EscHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EscHtml = s
End Function

'====================================================
' JSON �ַ���ת��
'====================================================
Private Function EscJsonStr(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\" & Chr$(34))
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscJsonStr = s
End Function

'====================================================
' ���򹤳�
'====================================================
Private Function MakeRE(ByVal sPat As String, _
                        Optional bGlobal As Boolean = True) As Object
    Set MakeRE = CreateObject("VBScript.RegExp")
    With MakeRE
        .Pattern = sPat
        .Global = bGlobal
        .IgnoreCase = False
        .Multiline = False
    End With
End Function

'====================================================
' Markdown �����жϺ���
'====================================================
Private Function HeadingLevel(ByVal ln As String) As Long
    Dim n As Long
    Do While n < 6 And n < Len(ln)
        If Mid$(ln, n + 1, 1) = "#" Then
            n = n + 1
        Else
            Exit Do
        End If
    Loop
    If n > 0 And (n >= Len(ln) Or Mid$(ln, n + 1, 1) = " ") Then
        HeadingLevel = n
    Else
        HeadingLevel = 0
    End If
End Function

Private Function IsHRule(ByVal ln As String) As Boolean
    Dim t As String
    t = Replace(Trim$(ln), " ", "")
    If Len(t) >= 3 Then
        IsHRule = (t = String$(Len(t), "-") Or _
                   t = String$(Len(t), "*") Or _
                   t = String$(Len(t), "_"))
    End If
End Function

Private Function IsULItem(ByVal ln As String) As Boolean
    Dim t As String
    t = LTrim$(ln)
    If Len(t) >= 2 Then
        IsULItem = (Left$(t, 2) = "- " Or Left$(t, 2) = "* " Or Left$(t, 2) = "+ ")
    End If
End Function

Private Function ULItemText(ByVal ln As String) As String
    ULItemText = Mid$(LTrim$(ln), 3)
End Function

Private Function IsOLItem(ByVal ln As String) As Boolean
    Dim re As Object
    Set re = MakeRE("^\s*\d+\.\s")
    IsOLItem = re.Test(ln)
End Function

Private Function OLItemText(ByVal ln As String) As String
    Dim re As Object, m As Object
    Set re = MakeRE("^\s*\d+\.\s(.*)")
    If re.Test(ln) Then
        Set m = re.Execute(ln)
        OLItemText = m(0).SubMatches(0)
    Else
        OLItemText = ln
    End If
End Function

Private Function IsTblRow(ByVal ln As String) As Boolean
    Dim t As String
    t = Trim$(ln)
    If Len(t) > 2 Then
        IsTblRow = (Left$(t, 1) = "|" And Right$(t, 1) = "|")
    End If
End Function

Private Function IsTblSep(ByVal ln As String) As Boolean
    Dim re As Object
    Set re = MakeRE("^\s*\|[\s\-:|]+\|\s*$")
    IsTblSep = re.Test(ln)
End Function

'====================================================
' UTF-8 �ļ���ȡ
'====================================================
Public Function ReadTextFile(ByVal sPath As String) As String
    On Error GoTo Err_Read
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile sPath
        ReadTextFile = .ReadText(-1)
        .Close
    End With
    Set stm = Nothing
    Exit Function
Err_Read:
    ReadTextFile = ""
    MsgBox "ReadTextFile: " & sPath & vbCrLf & Err.Description, vbExclamation
End Function

'====================================================
' �жϴ����Ƿ����
'====================================================
Private Function FormExists(ByVal sName As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = sName Then
            FormExists = True
            Exit Function
        End If
    Next obj
    FormExists = False
End Function


'############################################################
'#                                                          #
'#   ���岿��: ��ʾ                                          #
'#                                                          #
'############################################################

'====================================================
' Markdown ��ʾ
'====================================================
Public Sub MarkdownDemo()
    Dim s As String

    s = "# Markdown Demo" & vbCrLf & vbCrLf
    s = s & "## Basic Formatting" & vbCrLf & vbCrLf
    s = s & "Support **bold**, *italic*, ***bold italic***, ~~strikethrough~~ and `inline code`." & vbCrLf & vbCrLf

    s = s & "## Lists" & vbCrLf & vbCrLf
    s = s & "Unordered:" & vbCrLf
    s = s & "- Item A" & vbCrLf
    s = s & "- Item B" & vbCrLf
    s = s & "- Item C" & vbCrLf & vbCrLf
    s = s & "Ordered:" & vbCrLf
    s = s & "1. Step One" & vbCrLf
    s = s & "2. Step Two" & vbCrLf
    s = s & "3. Step Three" & vbCrLf & vbCrLf

    s = s & "## Blockquote" & vbCrLf & vbCrLf
    s = s & "> This is a blockquote." & vbCrLf
    s = s & "> Multiple lines." & vbCrLf & vbCrLf

    s = s & "## Code" & vbCrLf & vbCrLf
    s = s & "```vba" & vbCrLf
    s = s & "Sub Hello()" & vbCrLf
    s = s & "    MsgBox ""Hello!""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "```" & vbCrLf & vbCrLf

    s = s & "## Table" & vbCrLf & vbCrLf
    s = s & "| Name  | Dept   | ID  |" & vbCrLf
    s = s & "|-------|--------|-----|" & vbCrLf
    s = s & "| Alice | R&D    | 001 |" & vbCrLf
    s = s & "| Bob   | Sales  | 002 |" & vbCrLf & vbCrLf

    s = s & "---" & vbCrLf & vbCrLf
    s = s & "Link: [Microsoft](https://www.microsoft.com)" & vbCrLf & vbCrLf
    s = s & "*Rendered by Module_AI_Markdown*"

    ShowMarkdown s, "Markdown Demo"
End Sub


