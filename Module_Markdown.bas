Option Compare Database
Option Explicit
 
'
' 前置条件:
'   1. 导入 JsonConverter 模块 (VBA-JSON by Tim Hall)
'   2. 工具 -> 引用 -> 勾选 "Microsoft Scripting Runtime"
'   3. 修改各 AI 提供商的 API Key
'
' 支持的 AI 模型:
'   DeepSeek / OpenAI / 通义千问 / 文心一言 / Kimi / GLM / Gemini / 豆包 / 腾讯混元 / 讯飞星火 / 自定义
'
' 快速开始:
'   在 VBA 立即窗口执行:
'       CreateAIForm          ' 自动创建 AI 问答窗体
'   然后在 Access 中打开窗体 frmAI, 选择模型, 输入问题, 点击 [提问]
'
'   其他可用:
'       ShowMarkdown "# 标题" & vbCrLf & "**粗体**"
'====================================================

' ---------- Win32 Sleep ----------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SendMessageA Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function SendMessageA Lib "user32" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const WM_VSCROLL As Long = &H115
Private Const SB_BOTTOM As Long = 7

' ---------- AI 提供商配置 ----------
' 请将下方 Key 替换为你自己的 API Key

' DeepSeek
Private Const DS_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const DS_URL   As String = "https://api.deepseek.com/chat/completions"
Private Const DS_FLASH_MODEL As String = "deepseek-v4-flash"
Private Const DS_PRO_MODEL   As String = "deepseek-v4-pro"

' 通义千问 (阿里云百炼)
Private Const QW_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const QW_URL   As String = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
Private Const QW_MODEL As String = "qwen-plus"

' 文心一言 (百度千帆)
Private Const WX_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const WX_URL   As String = "https://qianfan.baidubce.com/v2/chat/completions"
Private Const WX_MODEL As String = "ernie-4.0-8k"

' Kimi (月之暗面)
Private Const KM_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const KM_URL   As String = "https://api.moonshot.cn/v1/chat/completions"
Private Const KM_MODEL As String = "moonshot-v1-8k"

' OpenAI
Private Const OA_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const OA_URL   As String = "https://api.openai.com/v1/chat/completions"
Private Const OA_GPT55_MODEL As String = "gpt-5.5"
Private Const OA_GPT54_MODEL As String = "gpt-5.4"

' 智谱清言 GLM (BigModel)
Private Const GLM_KEY  As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const GLM_URL  As String = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
Private Const GLM_FLASH_MODEL As String = "glm-4-flash"
Private Const GLM_PLUS_MODEL  As String = "glm-4-plus"

' Gemini (OpenAI 兼容接口)
Private Const GM_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const GM_URL   As String = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
Private Const GM_FLASH_MODEL As String = "gemini-1.5-flash"
Private Const GM_PRO_MODEL   As String = "gemini-1.5-pro"

' 豆包 (OpenAI 兼容接口)
Private Const DB_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const DB_URL   As String = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
Private Const DB_MODEL As String = "doubao-pro-32k"

' 腾讯混元 (OpenAI 兼容接口)
Private Const HY_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const HY_URL   As String = "https://api.hunyuan.cloud.tencent.com/v1/chat/completions"
Private Const HY_MODEL As String = "hunyuan-turbos-latest"

' 讯飞星火 (OpenAI 兼容接口)
Private Const XF_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const XF_URL   As String = "https://spark-api-open.xf-yun.com/v1/chat/completions"
Private Const XF_MODEL As String = "generalv3.5"

Private Const AI_FORM   As String = "frmAI"
Private Const AI_WEB_FORM As String = "frmAIWeb"
Private Const MD_FORM   As String = "frmMarkdownViewer"
Private Const TXT_MD    As String = "txtMarkdown"

' ---------- 对话历史记录 ----------
Private m_colHistory As Collection
Private m_sLastAnswer As String
Private m_sSessionId As String

' 当前会话在 txtAnswer 中累积的富文本 HTML(对话气泡)
Private m_sChatHtml As String
Private m_sStreamingAnswer As String

Private Const HISTORY_TABLE As String = "tblChatHistory"
Private Const HISTORY_FORM As String = "frmChatHistory"


'############################################################
'#                                                          #
'#   第一部分: Markdown -> 富文本 HTML                       #
'#                                                          #
'############################################################

'====================================================
' 核心: Markdown -> Access 富文本 HTML
' Access 支持: <b> <i> <u> <p> <br> <font> <ul> <ol> <li>
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

        ' ---- 代码块 ``` ----
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

        ' ---- 空行 ----
        If Len(Trim$(ln)) = 0 Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            GoTo NxtLine
        End If

        ' ---- 水平线 ----
        If IsHRule(ln) Then
            If inUL Then: out = out & "</ul>": inUL = False
            If inOL Then: out = out & "</ol>": inOL = False
            out = out & "<p><font color=""#cccccc"">" & String$(40, "-") & "</font></p>"
            GoTo NxtLine
        End If

        ' ---- 标题 # ~ ###### ----
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

        ' ---- 引用 > ----
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

        ' ---- 无序列表 ----
        If IsULItem(ln) Then
            If inOL Then: out = out & "</ol>": inOL = False
            If Not inUL Then: out = out & "<ul>": inUL = True
            out = out & "<li>" & FmtInline(ULItemText(ln)) & "</li>"
            GoTo NxtLine
        Else
            If inUL Then: out = out & "</ul>": inUL = False
        End If

        ' ---- 有序列表 ----
        If IsOLItem(ln) Then
            If inUL Then: out = out & "</ul>": inUL = False
            If Not inOL Then: out = out & "<ol>": inOL = True
            out = out & "<li>" & FmtInline(OLItemText(ln)) & "</li>"
            GoTo NxtLine
        Else
            If inOL Then: out = out & "</ol>": inOL = False
        End If

        ' ---- 表格 ----
        If IsTblRow(ln) Then
            If Not IsTblSep(ln) Then
                out = out & "<p><font face=""Consolas"" size=""2"">" & EscHtml(ln) & "</font></p>"
            End If
            GoTo NxtLine
        End If

        ' ---- 普通段落 ----
        out = out & "<p>" & FmtInline(ln) & "</p>"

NxtLine:
    Next i

    If inUL Then out = out & "</ul>"
    If inOL Then out = out & "</ol>"

    MarkdownToRichText = out
End Function

'====================================================
' 行内格式 (粗体/斜体/代码/链接/图片/删除线)
'====================================================
Private Function FmtInline(ByVal s As String) As String
    Dim re As Object

    s = EscHtml(s)

    ' `code`
    Set re = MakeRE("`([^`]+)`")
    s = re.Replace(s, "<font face=""Consolas"" color=""#c7254e"">$1</font>")

    ' ![alt](url) - 图片 (必须在链接之前)
    Set re = MakeRE("!\[([^\]]*)\]\(([^)]+)\)")
    s = re.Replace(s, "<font color=""#999999"">[img: $1]</font>")

    ' [text](url) - 链接
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
'#   第二部分: AI API 调用 (多模型支持)                      #
'#   支持: DeepSeek/OpenAI/通义千问/文心一言/Kimi/GLM/Gemini/豆包等 OpenAI 兼容接口 #
'#   方案A: curl 子进程真流式 (Windows 10 1803+)            #
'#   方案B: 同步请求 + 打字机效果 (兜底)                     #
'#                                                          #
'############################################################

'====================================================
' 根据提供商名称返回 API 配置
'====================================================
Private Function GetProviderRowSource() As String
    Dim vProviders As Variant
    Dim i As Long
    Dim sRows As String

    vProviders = Array( _
        "DeepSeek Flash", "DeepSeek Pro", _
        "通义千问 Plus", "通义千问", _
        "文心一言", "Kimi", _
        "OpenAI GPT-5.5", "OpenAI GPT-5.4", _
        "GLM Flash", "GLM Plus", _
        "Gemini Flash", "Gemini Pro", _
        "豆包", "腾讯混元", "讯飞星火", _
        "自定义")

    For i = LBound(vProviders) To UBound(vProviders)
        If Len(sRows) > 0 Then sRows = sRows & ";"
        sRows = sRows & """" & CStr(vProviders(i)) & """"
    Next i
    GetProviderRowSource = sRows
End Function

Private Sub GetProviderConfig(ByVal sProvider As String, _
                              ByRef sUrl As String, _
                              ByRef sKey As String, _
                              ByRef sModel As String)
    Select Case sProvider
        Case "DeepSeek Flash"
            sUrl = DS_URL: sKey = DS_KEY: sModel = DS_FLASH_MODEL
        Case "DeepSeek Pro", "DeepSeek"
            sUrl = DS_URL: sKey = DS_KEY: sModel = DS_PRO_MODEL
        Case "通义千问", "通义千问 Plus"
            sUrl = QW_URL: sKey = QW_KEY: sModel = QW_MODEL
        Case "文心一言"
            sUrl = WX_URL: sKey = WX_KEY: sModel = WX_MODEL
        Case "Kimi"
            sUrl = KM_URL: sKey = KM_KEY: sModel = KM_MODEL
        Case "OpenAI GPT-5.5"
            sUrl = OA_URL: sKey = OA_KEY: sModel = OA_GPT55_MODEL
        Case "OpenAI GPT-5.4"
            sUrl = OA_URL: sKey = OA_KEY: sModel = OA_GPT54_MODEL
        Case "GLM Flash"
            sUrl = GLM_URL: sKey = GLM_KEY: sModel = GLM_FLASH_MODEL
        Case "GLM Plus"
            sUrl = GLM_URL: sKey = GLM_KEY: sModel = GLM_PLUS_MODEL
        Case "Gemini Flash"
            sUrl = GM_URL: sKey = GM_KEY: sModel = GM_FLASH_MODEL
        Case "Gemini Pro"
            sUrl = GM_URL: sKey = GM_KEY: sModel = GM_PRO_MODEL
        Case "豆包"
            sUrl = DB_URL: sKey = DB_KEY: sModel = DB_MODEL
        Case "腾讯混元"
            sUrl = HY_URL: sKey = HY_KEY: sModel = HY_MODEL
        Case "讯飞星火"
            sUrl = XF_URL: sKey = XF_KEY: sModel = XF_MODEL
        Case "自定义"
            Dim frmC As Form
            Set frmC = Screen.ActiveForm
            sUrl = Nz(frmC!txtCustomUrl, "")
            sKey = Nz(frmC!txtCustomKey, "")
            sModel = Nz(frmC!txtCustomModel, "")
        Case Else  ' DeepSeek Pro (默认)
            sUrl = DS_URL: sKey = DS_KEY: sModel = DS_PRO_MODEL
    End Select
End Sub

'====================================================
' 对话历史管理
'====================================================
Private Sub InitHistory()
    If m_colHistory Is Nothing Then
        Set m_colHistory = New Collection
    End If
    If Len(m_sSessionId) = 0 Then
        m_sSessionId = NewSessionId()
    End If
End Sub

Public Sub ClearHistory()
    Set m_colHistory = New Collection
    m_sLastAnswer = ""
    m_sSessionId = NewSessionId()
    m_sChatHtml = ""
    m_sStreamingAnswer = ""
End Sub

'====================================================
' 对话气泡 (Access 富文本 HTML)
'   仅支持 <div align> / <font color> / <b> 等有限标签,
'   无法设置背景色与圆角, 因此用 "右对齐+主题色" 代表用户,
'   "左对齐+灰色" 代表 AI, 模拟网页聊天列表效果
'====================================================
Private Function HtmlEscapeText(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlEscapeText = s
End Function

Private Function TextToRtBr(ByVal s As String) As String
    s = HtmlEscapeText(s)
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "<br>")
    TextToRtBr = s
End Function

Private Function BuildUserBubbleHtml(ByVal sText As String) As String
    BuildUserBubbleHtml = _
        "<div align=""right""><font color=""#4E6CFE"" face=""Microsoft YaHei""><b>我</b></font></div>" & _
        "<div align=""right""><font color=""#1D1E20"" face=""Microsoft YaHei"">" & TextToRtBr(sText) & "</font></div>" & _
        "<div>&nbsp;</div>"
End Function

Private Function BuildAiBubbleHtml(ByVal sMarkdown As String) As String
    BuildAiBubbleHtml = _
        "<div align=""left""><font color=""#868E99"" face=""Microsoft YaHei""><b>AI</b></font></div>" & _
        "<div align=""left""><font color=""#1D1E20"" face=""Microsoft YaHei"">" & MarkdownToRichText(sMarkdown) & "</font></div>" & _
        "<div>&nbsp;</div>"
End Function

Private Function BuildAiStreamingBubbleHtml(ByVal sText As String, ByVal bCursor As Boolean) As String
    Dim s As String
    s = TextToRtBr(sText)
    If bCursor Then s = s & "<font color=""#4E6CFE"">&#9612;</font>"
    BuildAiStreamingBubbleHtml = _
        "<div align=""left""><font color=""#868E99"" face=""Microsoft YaHei""><b>AI</b></font></div>" & _
        "<div align=""left""><font color=""#1D1E20"" face=""Microsoft YaHei"">" & s & "</font></div>" & _
        "<div>&nbsp;</div>"
End Function

'====================================================
' 将 txtAnswer 滚动到最底部 (显示最新消息)
'   思路: SetFocus 后用 GetFocus() 拿到 Access 内部编辑控件 hWnd,
'         再发 WM_VSCROLL(SB_BOTTOM) 把滚动条拉到底.
'         参数 bKeepFocus=True 时保持焦点在 txtAnswer,
'         这样用户可以直接用鼠标滚轮上下滚.
'====================================================
Private Sub ScrollAnswerToEnd(frm As Form, Optional ByVal bKeepFocus As Boolean = True)
    On Error Resume Next
    If HasControl(frm, "wbChat") Then
        RefreshAlternateChatView frm
        Exit Sub
    End If

    Dim ctlPrev As Control
    Set ctlPrev = Screen.ActiveControl

    PrepareAnswerBox frm
    frm!txtAnswer.SetFocus
#If VBA7 Then
    Dim hEdit As LongPtr
#Else
    Dim hEdit As Long
#End If
    hEdit = GetFocus()
    If hEdit <> 0 Then
        SendMessageA hEdit, WM_VSCROLL, SB_BOTTOM, 0
    End If

    If Not bKeepFocus Then
        If Not ctlPrev Is Nothing Then
            If ctlPrev.Name <> "txtAnswer" Then ctlPrev.SetFocus
        Else
            frm!txtQ.SetFocus
        End If
    End If
End Sub

Private Sub PrepareAnswerBox(frm As Form)
    On Error Resume Next
    With frm!txtAnswer
        .Enabled = True
        .Locked = False
        .TabStop = True
        .ScrollBars = 2
        .TextFormat = acTextFormatHTMLRichText
    End With
End Sub

Private Function HasControl(frm As Form, ByVal sControlName As String) As Boolean
    On Error GoTo NotFound
    Dim ctl As Control
    Set ctl = frm.Controls(sControlName)
    HasControl = True
    Exit Function
NotFound:
    HasControl = False
End Function

Private Sub RefreshAlternateChatView(frm As Form)
    On Error Resume Next
    If HasControl(frm, "wbChat") Then
        RenderWebChat frm
    End If
End Sub

'====================================================
' 从当前 m_colHistory 重建全量气泡 HTML (加载历史会话时用)
'====================================================
Private Sub RebuildChatHtmlFromHistory()
    Dim i As Long
    Dim sRole As String, sContent As String
    m_sChatHtml = ""
    If m_colHistory Is Nothing Then Exit Sub
    For i = 1 To m_colHistory.Count
        sRole = CStr(m_colHistory(i)("role"))
        sContent = CStr(m_colHistory(i)("content"))
        If sRole = "user" Then
            m_sChatHtml = m_sChatHtml & BuildUserBubbleHtml(sContent)
        ElseIf sRole = "assistant" Then
            m_sChatHtml = m_sChatHtml & BuildAiBubbleHtml(sContent)
        End If
    Next i
End Sub

Private Function GetSystemPromptFromForm(frm As Form) As String
    On Error Resume Next
    GetSystemPromptFromForm = Trim$(Nz(frm!txtSystemPrompt, ""))
    If Err.Number = 0 Then SaveSystemPrompt GetSystemPromptFromForm
End Function

Private Function GetSavedSystemPrompt() As String
    GetSavedSystemPrompt = GetSetting("AccessAI", "Settings", "SystemPrompt", "")
End Function

Private Sub SaveSystemPrompt(ByVal sSystemPrompt As String)
    SaveSetting "AccessAI", "Settings", "SystemPrompt", sSystemPrompt
End Sub

Private Function GetReasoningEffortFromForm(frm As Form) As String
    On Error Resume Next
    GetReasoningEffortFromForm = Trim$(Nz(frm!cboReasoningEffort, ""))
    If Err.Number = 0 Then SaveReasoningEffort GetReasoningEffortFromForm
End Function

Private Function GetSavedReasoningEffort() As String
    GetSavedReasoningEffort = GetSetting("AccessAI", "Settings", "ReasoningEffort", "默认")
End Function

Private Sub SaveReasoningEffort(ByVal sReasoningEffort As String)
    If Len(Trim$(sReasoningEffort)) = 0 Then sReasoningEffort = "默认"
    SaveSetting "AccessAI", "Settings", "ReasoningEffort", sReasoningEffort
End Sub

Private Function NormalizeReasoningEffort(ByVal sReasoningEffort As String) As String
    Select Case LCase$(Trim$(sReasoningEffort))
        Case "low", "medium", "high", "xhigh"
            NormalizeReasoningEffort = LCase$(Trim$(sReasoningEffort))
        Case Else
            NormalizeReasoningEffort = ""
    End Select
End Function

'====================================================
' 方案A: WebBrowser HTML 对话窗口渲染
'====================================================
Private Function WebHtmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    WebHtmlEscape = s
End Function

Private Function TextToWebHtml(ByVal s As String) As String
    s = WebHtmlEscape(s)
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "<br>")
    TextToWebHtml = s
End Function

Private Function MarkdownToWebHtml(ByVal sMd As String) As String
    MarkdownToWebHtml = MarkdownToRichText(sMd)
End Function

Private Function BuildWebBubble(ByVal sRole As String, ByVal sContent As String, Optional ByVal bStreaming As Boolean = False) As String
    Dim sBody As String

    If sRole = "user" Then
        sBody = TextToWebHtml(sContent)
    Else
        If bStreaming Then
            sBody = TextToWebHtml(sContent) & "<span class='cursor'></span>"
        Else
            sBody = MarkdownToWebHtml(sContent)
        End If
    End If

    If sRole = "user" Then
        BuildWebBubble = _
            "<table class='msgRow' width='100%' cellpadding='0' cellspacing='0' border='0' style='margin:0 0 14px 0;'>" & _
            "<tr><td width='24%'>&nbsp;</td><td width='76%' align='right'>" & _
            "<table class='userWrap' width='100%' cellpadding='0' cellspacing='0' border='0' style='background:#f2f8ee;border-right:4px solid #61b875;padding:10px 12px;'>" & _
            "<tr><td align='right' valign='top'>" & _
            "<table class='bubble userBubble' cellpadding='0' cellspacing='0' border='0' align='right' style='background:#2f7dff;border:1px solid #2f7dff;color:#ffffff;font-family:Microsoft YaHei,Arial;font-size:14px;line-height:1.65;text-align:left;'><tr><td style='padding:12px 15px;color:#ffffff;'>" & sBody & "</td></tr></table>" & _
            "</td><td width='44' align='right' valign='top'><div class='avatar userAvatar' style='width:34px;height:34px;line-height:34px;text-align:center;font-size:12px;font-weight:bold;background:#61b875;color:#ffffff;font-family:Microsoft YaHei,Arial;'>我</div></td></tr>" & _
            "</table></td></tr></table>"
    Else
        BuildWebBubble = _
            "<table class='msgRow' width='100%' cellpadding='0' cellspacing='0' border='0' style='margin:0 0 14px 0;'>" & _
            "<tr><td width='76%' align='left'>" & _
            "<table class='aiWrap' width='100%' cellpadding='0' cellspacing='0' border='0' style='background:#f7fbff;border-left:4px solid #6aa6ff;padding:10px 12px;'>" & _
            "<tr><td width='44' align='left' valign='top'><div class='avatar aiAvatar' style='width:34px;height:34px;line-height:34px;text-align:center;font-size:12px;font-weight:bold;background:#dcecff;color:#255f9e;font-family:Microsoft YaHei,Arial;'>AI</div></td>" & _
            "<td align='left' valign='top'><table class='bubble aiBubble' cellpadding='0' cellspacing='0' border='0' style='background:#ffffff;border:1px solid #dbe7f5;color:#1d1e20;font-family:Microsoft YaHei,Arial;font-size:14px;line-height:1.65;text-align:left;'><tr><td style='padding:12px 15px;color:#1d1e20;'>" & sBody & "</td></tr></table></td></tr>" & _
            "</table></td><td width='24%'>&nbsp;</td></tr></table>"
    End If
End Function

Private Function BuildWebChatBody(ByVal frm As Form) As String
    On Error Resume Next
    Dim i As Long
    Dim sBody As String
    Dim sRole As String
    Dim sContent As String

    If Not m_colHistory Is Nothing Then
        For i = 1 To m_colHistory.Count
            sRole = CStr(m_colHistory(i)("role"))
            sContent = CStr(m_colHistory(i)("content"))
            If sRole = "user" Or sRole = "assistant" Then
                sBody = sBody & BuildWebBubble(sRole, sContent)
            End If
        Next i
    End If

    If Len(m_sStreamingAnswer) > 0 Then
        sBody = sBody & BuildWebBubble("assistant", m_sStreamingAnswer, (Len(m_sLastAnswer) = 0))
    End If

    If Len(sBody) = 0 Then
        sBody = "<div class='empty'>开始一次对话，AI 的回复会显示在这里。</div>"
    End If
    BuildWebChatBody = sBody
End Function

Private Function ExtractStreamingTextFromRichHtml(ByVal sHtml As String) As String
    On Error Resume Next
    Dim s As String
    Dim p As Long
    s = sHtml
    p = InStrRev(s, "<b>AI</b>")
    If p > 0 Then s = Mid$(s, p + Len("<b>AI</b>"))
    s = Replace(s, "&#9612;", "")
    s = Replace(s, "<br>", vbCrLf)
    s = Replace(s, "<div>&nbsp;</div>", "")
    Dim re As Object
    Set re = MakeRE("<[^>]+>")
    s = re.Replace(s, "")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&amp;", "&")
    s = Trim$(s)
    ExtractStreamingTextFromRichHtml = s
End Function

Private Function BuildWebChatDocument(ByVal sBodyHtml As String) As String
    BuildWebChatDocument = "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'>" & _
        "<meta charset='utf-8'><style>" & _
        "html,body{height:100%;}body{margin:0;padding:18px 22px;font-family:'Microsoft YaHei',Segoe UI,Arial,sans-serif;background:#e9eef5;color:#1d1e20;font-size:14px;line-height:1.65;}" & _
        ".msgRow{margin:0 0 14px 0;}.aiWrap{background:#f7fbff;border-left:4px solid #6aa6ff;}.userWrap{background:#f2f8ee;border-right:4px solid #61b875;}" & _
        ".aiWrap,.userWrap{padding:10px 12px;}.avatar{width:34px;height:34px;line-height:34px;text-align:center;font-size:12px;font-weight:bold;font-family:'Microsoft YaHei',Arial,sans-serif;}" & _
        ".aiAvatar{background:#dcecff;color:#255f9e;}.userAvatar{background:#61b875;color:#ffffff;}" & _
        ".bubble{font-family:'Microsoft YaHei',Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.65;text-align:left;word-wrap:break-word;}.bubble td{padding:12px 15px;}.aiBubble{background:#ffffff;border:1px solid #dbe7f5;color:#1d1e20;}.userBubble{background:#2f7dff;border:1px solid #2f7dff;color:#ffffff;}" & _
        ".bubble p{margin:5px 0 10px}.bubble ul,.bubble ol{margin-top:6px;margin-bottom:10px;padding-left:22px}.bubble code,.bubble pre,.bubble .code{font-family:Consolas,monospace}.bubble a{color:#0366d6}.userBubble font,.userBubble a{color:#fff!important}.empty{padding-top:120px;text-align:center;color:#7d8796;font-size:13px}.cursor{display:inline-block;width:8px;height:16px;margin-left:2px;background:#2f7dff;}" & _
        "</style></head><body>" & sBodyHtml & _
        "<script>window.scrollTo(0,document.body.scrollHeight);</script></body></html>"
End Function

Private Function TryCreateWebBrowserControl(frm As Form, ByRef ctlOut As Control) As Boolean
    Dim vProgIds As Variant
    Dim i As Long
    Dim ctl As Control

    vProgIds = Array("Shell.Explorer.2", "Shell.Explorer", "Microsoft Web Browser", "WebBrowser.WebBrowser.1", "{8856F961-340A-11D0-A96B-00C04FD705A2}")
    For i = LBound(vProgIds) To UBound(vProgIds)
        Err.Clear
        On Error Resume Next
        Set ctl = CreateControl(frm.Name, acCustomControl, acDetail, , CStr(vProgIds(i)), 500, 2080, 13400, 6360)
        If Err.Number = 0 And Not ctl Is Nothing Then
            ctl.Name = "wbChat"
            Set ctlOut = ctl
            TryCreateWebBrowserControl = True
            On Error GoTo 0
            Exit Function
        End If
        Set ctl = Nothing
        On Error GoTo 0
    Next i
End Function

Private Sub RenderWebChat(frm As Form)
    On Error Resume Next
    Dim wb As Object
    Set wb = frm!wbChat.Object
    If wb Is Nothing Then Exit Sub

    If wb.LocationURL = "" Then
        wb.Navigate "about:blank"
        DoEvents
    End If

    Dim sHtml As String
    sHtml = BuildWebChatDocument(BuildWebChatBody(frm))
    wb.Document.Open
    wb.Document.Write sHtml
    wb.Document.Close
    wb.Document.parentWindow.scrollTo 0, wb.Document.body.scrollHeight
End Sub

'====================================================
' 当前数据库对象分析辅助
'====================================================
Private Function BracketName(ByVal sName As String) As String
    BracketName = "[" & Replace(sName, "]", "]]") & "]"
End Function

Private Function MdCell(ByVal v As Variant, Optional ByVal lMaxLen As Long = 120) As String
    Dim s As String
    If IsNull(v) Then
        MdCell = "(Null)"
        Exit Function
    End If
    If IsDate(v) Then
        s = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
    Else
        s = CStr(v)
    End If
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "|", "\|")
    If Len(s) > lMaxLen Then s = Left$(s, lMaxLen) & "..."
    MdCell = s
End Function

Private Function DaoTypeName(ByVal lType As Long) As String
    Select Case lType
        Case dbBoolean: DaoTypeName = "Yes/No"
        Case dbByte: DaoTypeName = "Byte"
        Case dbInteger: DaoTypeName = "Integer"
        Case dbLong: DaoTypeName = "Long"
        Case dbCurrency: DaoTypeName = "Currency"
        Case dbSingle: DaoTypeName = "Single"
        Case dbDouble: DaoTypeName = "Double"
        Case dbDate: DaoTypeName = "Date/Time"
        Case dbText: DaoTypeName = "Short Text"
        Case dbLongBinary: DaoTypeName = "OLE/Object"
        Case dbMemo: DaoTypeName = "Long Text"
        Case dbGUID: DaoTypeName = "GUID"
        Case 16: DaoTypeName = "BigInt"
        Case 101: DaoTypeName = "Attachment/Complex"
        Case 102 To 109: DaoTypeName = "Complex"
        Case Else
            DaoTypeName = "Type " & CStr(lType)
    End Select
End Function

Private Function IsUserTableName(ByVal sName As String) As Boolean
    IsUserTableName = (Left$(sName, 4) <> "MSys" And Left$(sName, 1) <> "~")
End Function

Private Function QuoteValueList(ByVal s As String) As String
    QuoteValueList = """" & Replace(Replace(s, """", "'"), ";", ",") & """"
End Function

Private Function GetDbObjectRowSource() As String
    On Error Resume Next
    Dim sRows As String
    Dim obj As AccessObject

    For Each obj In CurrentData.AllTables
        If IsUserTableName(obj.Name) Then
            If Len(sRows) > 0 Then sRows = sRows & ";"
            sRows = sRows & QuoteValueList("表: " & obj.Name)
        End If
    Next obj

    For Each obj In CurrentData.AllQueries
        If Len(sRows) > 0 Then sRows = sRows & ";"
        sRows = sRows & QuoteValueList("查询: " & obj.Name)
    Next obj

    GetDbObjectRowSource = sRows
End Function

Private Function ParseDbObjectName(ByVal sDisplay As String) As String
    If Left$(sDisplay, 3) = "表: " Then
        ParseDbObjectName = Mid$(sDisplay, 4)
    ElseIf Left$(sDisplay, 4) = "查询: " Then
        ParseDbObjectName = Mid$(sDisplay, 5)
    Else
        ParseDbObjectName = sDisplay
    End If
End Function

Private Function ParseDbObjectKind(ByVal sDisplay As String) As String
    If Left$(sDisplay, 3) = "表: " Then
        ParseDbObjectKind = "Table"
    ElseIf Left$(sDisplay, 4) = "查询: " Then
        ParseDbObjectKind = "Query"
    Else
        ParseDbObjectKind = "Table/Query"
    End If
End Function

Private Function BuildDbObjectContext(ByVal sDisplayName As String, Optional ByVal lTopN As Long = 30) As String
    On Error GoTo ErrHandler
    Dim sObjectName As String
    Dim sKind As String
    Dim db As DAO.Database
    Dim rsSchema As DAO.Recordset
    Dim rsSample As DAO.Recordset
    Dim rsCount As DAO.Recordset
    Dim fld As DAO.Field
    Dim sSqlName As String
    Dim sOut As String
    Dim lCount As Long
    Dim i As Long
    Dim lRow As Long

    sObjectName = ParseDbObjectName(sDisplayName)
    sKind = ParseDbObjectKind(sDisplayName)
    sSqlName = BracketName(sObjectName)
    Set db = CurrentDb

    Set rsCount = db.OpenRecordset("SELECT Count(*) AS Cnt FROM " & sSqlName, dbOpenSnapshot)
    If Not rsCount.EOF Then lCount = CLng(Nz(rsCount!Cnt, 0))
    rsCount.Close

    Set rsSchema = db.OpenRecordset("SELECT * FROM " & sSqlName & " WHERE 1=0", dbOpenSnapshot)
    Set rsSample = db.OpenRecordset("SELECT TOP " & CStr(lTopN) & " * FROM " & sSqlName, dbOpenSnapshot)

    sOut = "## 数据对象" & vbCrLf & _
           "- 名称: " & sObjectName & vbCrLf & _
           "- 类型: " & sKind & vbCrLf & _
           "- 记录数: " & CStr(lCount) & vbCrLf & _
           "- 样例行数: " & CStr(lTopN) & vbCrLf & vbCrLf

    sOut = sOut & "## 字段结构" & vbCrLf & "| 字段 | 类型 | 大小 |" & vbCrLf & "|---|---|---:|" & vbCrLf
    For Each fld In rsSchema.Fields
        sOut = sOut & "| " & MdCell(fld.Name) & " | " & DaoTypeName(fld.Type) & " | " & CStr(fld.Size) & " |" & vbCrLf
    Next fld

    sOut = sOut & vbCrLf & "## 样例数据" & vbCrLf
    If rsSchema.Fields.Count = 0 Then
        sOut = sOut & "(无字段)" & vbCrLf
    Else
        sOut = sOut & "|"
        For i = 0 To rsSchema.Fields.Count - 1
            sOut = sOut & " " & MdCell(rsSchema.Fields(i).Name) & " |"
        Next i
        sOut = sOut & vbCrLf & "|"
        For i = 0 To rsSchema.Fields.Count - 1
            sOut = sOut & "---|"
        Next i
        sOut = sOut & vbCrLf

        Do While Not rsSample.EOF And lRow < lTopN
            sOut = sOut & "|"
            For i = 0 To rsSample.Fields.Count - 1
                sOut = sOut & " " & MdCell(rsSample.Fields(i).Value) & " |"
            Next i
            sOut = sOut & vbCrLf
            lRow = lRow + 1
            rsSample.MoveNext
        Loop
        If lRow = 0 Then sOut = sOut & "(无样例数据)" & vbCrLf
    End If

    rsSchema.Close
    rsSample.Close

    BuildDbObjectContext = sOut
    Exit Function

ErrHandler:
    BuildDbObjectContext = "## 数据对象读取失败" & vbCrLf & _
                           "- 对象: " & sDisplayName & vbCrLf & _
                           "- 错误: " & Err.Description & vbCrLf
End Function

'====================================================
' 确保历史记录表存在
'====================================================
Private Sub EnsureHistoryTable()
    Dim db As Object ' DAO.Database
    Dim td As Object ' DAO.TableDef
    Dim fld As Object ' DAO.Field
    Dim idx As Object ' DAO.Index

    Set db = CurrentDb

    ' 检查表是否已存在
    Dim bExists As Boolean
    Dim tbl As AccessObject
    bExists = False
    For Each tbl In CurrentData.AllTables
        If tbl.Name = HISTORY_TABLE Then
            bExists = True
            Exit For
        End If
    Next tbl
    If bExists Then Exit Sub

    ' 创建表
    Set td = db.CreateTableDef(HISTORY_TABLE)

    Set fld = td.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    td.Fields.Append fld

    td.Fields.Append td.CreateField("SessionID", dbText, 50)
    td.Fields.Append td.CreateField("Provider", dbText, 50)
    td.Fields.Append td.CreateField("Role", dbText, 20)
    td.Fields.Append td.CreateField("Content", dbMemo)
    td.Fields.Append td.CreateField("CreatedAt", dbDate)

    db.TableDefs.Append td

    ' 主键
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Fields.Append idx.CreateField("ID")
    td.Indexes.Append idx

    ' SessionID 索引
    Set idx = td.CreateIndex("idxSession")
    idx.Fields.Append idx.CreateField("SessionID")
    td.Indexes.Append idx

    db.TableDefs.Refresh
End Sub

'====================================================
' 生成新的会话 ID
'====================================================
Private Function NewSessionId() As String
    Randomize
    NewSessionId = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 10000))
End Function

'====================================================
' 保存消息到数据库
'====================================================
Private Sub SaveMessageToDb(ByVal sSessionId As String, _
                            ByVal sProvider As String, _
                            ByVal sRole As String, _
                            ByVal sContent As String)
    On Error Resume Next
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(HISTORY_TABLE, dbOpenDynaset)
    rs.AddNew
    rs!SessionID = sSessionId
    rs!Provider = sProvider
    rs!Role = sRole
    rs!Content = sContent
    rs!CreatedAt = Now
    rs.Update
    rs.Close
    Set rs = Nothing
End Sub

'====================================================
' 按钮事件: 发送
'====================================================
Public Function btnAsk_Click()
    Askai
End Function

'====================================================
' 按钮事件: 新对话
'====================================================
Public Function btnNewChat_Click()
    ClearHistory
    Dim frm As Form
    Set frm = Screen.ActiveForm
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = ""
    RefreshAlternateChatView frm
    frm!lblMsg.Caption = "已开始新对话。选择模型，输入问题后点击发送"
End Function

'====================================================
' 模型切换事件: 显示/隐藏自定义端点字段
'====================================================
Public Function cboProvider_AfterUpdate()
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    Dim bCustom As Boolean
    bCustom = (Nz(frm!cboProvider, "") = "自定义")
    frm!rectCustomBg.Visible = bCustom
    frm!lblCustomUrl.Visible = bCustom
    frm!txtCustomUrl.Visible = bCustom
    frm!lblCustomKey.Visible = bCustom
    frm!txtCustomKey.Visible = bCustom
    frm!lblCustomModel.Visible = bCustom
    frm!txtCustomModel.Visible = bCustom
End Function

'====================================================
' 系统提示词变更事件: 保存配置
'====================================================
Public Function txtSystemPrompt_AfterUpdate()
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    SaveSystemPrompt Trim$(Nz(frm!txtSystemPrompt, ""))
End Function

'====================================================
' 思考强度变更事件: 保存配置
'====================================================
Public Function cboReasoningEffort_AfterUpdate()
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    SaveReasoningEffort Nz(frm!cboReasoningEffort, "默认")
End Function

'====================================================
' 数据对象下拉框获取焦点: 刷新表/查询列表
'====================================================
Public Function cboDbObject_GotFocus()
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    frm!cboDbObject.RowSource = GetDbObjectRowSource()
    frm!cboDbObject.Requery
End Function

'====================================================
' 按钮事件: 分析当前数据库中的表/查询
'====================================================
Public Function btnAnalyzeData_Click()
    On Error GoTo ErrHandler
    Dim frm As Form
    Set frm = Screen.ActiveForm

    Dim sDisplayName As String
    sDisplayName = Nz(frm!cboDbObject, "")
    If Len(Trim$(sDisplayName)) = 0 Then
        MsgBox "请先选择一个表或查询。", vbInformation
        Exit Function
    End If

    Dim sQuestion As String
    sQuestion = Trim$(Nz(frm!txtQ, ""))
    If Len(sQuestion) = 0 Then
        sQuestion = "请分析这个数据对象的业务含义、关键字段、数据质量问题、可分析方向，并给出建议的 Access SQL 查询。"
    End If

    frm!lblMsg.Caption = "正在读取数据库对象..."
    frm.Repaint

    Dim sContext As String
    sContext = BuildDbObjectContext(sDisplayName, 30)
    If Left$(sContext, Len("## 数据对象读取失败")) = "## 数据对象读取失败" Then
        frm!txtQ.Value = sContext
        frm!lblMsg.Caption = "数据库对象读取失败。"
        MsgBox "读取表/查询失败，请确认该对象不是参数查询、操作查询或受权限限制。", vbExclamation
        Exit Function
    End If

    frm!txtQ.Value = sQuestion & vbCrLf & vbCrLf & _
                     "下面是当前 Access 数据库对象的结构和样例数据，请基于这些内容分析，不要假设未提供的数据。" & vbCrLf & vbCrLf & _
                     sContext
    Askai
    Exit Function

ErrHandler:
    MsgBox "btnAnalyzeData_Click: " & Err.Description, vbExclamation
End Function

'====================================================
' 按钮事件: 历史记录
'====================================================
Public Function btnHistory_Click()
    ShowChatHistory
End Function

'====================================================
' 显示历史对话记录
'====================================================
Public Sub ShowChatHistory()
    On Error GoTo ErrHandler

    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Count(*) AS Cnt FROM " & HISTORY_TABLE, dbOpenSnapshot)
    If rs!Cnt = 0 Then
        rs.Close
        MsgBox "暂无历史对话记录。", vbInformation
        Exit Sub
    End If
    rs.Close

    If Not FormExists(HISTORY_FORM) Then
        CreateHistoryForm
    End If

    DoCmd.OpenForm HISTORY_FORM, acNormal
    RefreshSessionList
    Exit Sub

ErrHandler:
    MsgBox "ShowChatHistory: " & Err.Description, vbExclamation
End Sub

'====================================================
' 刷新会话列表
'====================================================
Private Sub RefreshSessionList()
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(HISTORY_FORM)

    Dim db As Object 'DAO.Database
    Dim rs As Object ' DAO.Recordset
    Set db = CurrentDb

    Dim sSQL As String
    sSQL = "SELECT TOP 50 t1.SessionID, t1.CreatedAt, t1.Provider, Left(t1.Content, 50) AS Preview " & _
           "FROM " & HISTORY_TABLE & " AS t1 " & _
           "WHERE t1.Role='user' AND t1.ID = " & _
           "(SELECT MIN(t2.ID) FROM " & HISTORY_TABLE & " AS t2 WHERE t2.SessionID = t1.SessionID) " & _
           "ORDER BY t1.CreatedAt DESC"
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)

    Dim sValueList As String
    sValueList = ""
    Do While Not rs.EOF
        Dim sDisplay As String
        sDisplay = Format(rs!CreatedAt, "yyyy/mm/dd hh:nn") & " [" & Nz(rs!Provider, "") & "] " & _
                   Nz(rs!Preview, "")
        sDisplay = Replace(sDisplay, """", "'")
        sDisplay = Replace(sDisplay, ";", ",")
        sDisplay = Replace(Replace(Replace(sDisplay, vbCrLf, " "), vbCr, " "), vbLf, " ")
        If Len(sValueList) > 0 Then sValueList = sValueList & ";"
        sValueList = sValueList & """" & rs!SessionID & """;""" & sDisplay & """"
        rs.MoveNext
    Loop
    rs.Close

    frm!cboSession.RowSource = sValueList
    frm!cboSession.Requery
End Sub

'====================================================
' 会话选择事件: 显示对话详情
'====================================================
Public Function cboSession_AfterUpdate()
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(HISTORY_FORM)

    Dim sSessId As String
    sSessId = Nz(frm!cboSession, "")
    If Len(sSessId) = 0 Then Exit Function

    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT Role, Content, CreatedAt FROM " & HISTORY_TABLE & _
        " WHERE SessionID='" & Replace(sSessId, "'", "''") & "'" & _
        " ORDER BY ID", dbOpenSnapshot)

    Dim sOut As String
    sOut = ""
    Do While Not rs.EOF
        If rs!Role = "user" Then
            sOut = sOut & "### " & Chr(9654) & " 用户 (" & Format(rs!CreatedAt, "hh:nn:ss") & ")" & vbLf & vbLf
            sOut = sOut & rs!Content & vbLf & vbLf
        Else
            sOut = sOut & "### AI 回答" & vbLf & vbLf
            sOut = sOut & rs!Content & vbLf & vbLf
            sOut = sOut & "---" & vbLf & vbLf
        End If
        rs.MoveNext
    Loop
    rs.Close

    frm!txtHistoryDetail.TextFormat = acTextFormatHTMLRichText
    frm!txtHistoryDetail.Value = MarkdownToRichText(sOut)
End Function

'====================================================
' 加载历史对话到当前会话
'====================================================
Public Function btnLoadSession_Click()
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(HISTORY_FORM)

    Dim sSessId As String
    sSessId = Nz(frm!cboSession, "")
    If Len(sSessId) = 0 Then
        MsgBox "请先选择一个会话。", vbInformation
        Exit Function
    End If

    ' 从数据库加载到内存
    Set m_colHistory = New Collection
    m_sSessionId = sSessId

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT Role, Content FROM " & HISTORY_TABLE & _
        " WHERE SessionID='" & Replace(sSessId, "'", "''") & "'" & _
        " ORDER BY ID", dbOpenSnapshot)

    Dim sLast As String
    sLast = ""
    Do While Not rs.EOF
        Dim oMsg As Object
        Set oMsg = CreateObject("Scripting.Dictionary")
        oMsg.Add "role", CStr(rs!Role)
        oMsg.Add "content", CStr(rs!Content)
        m_colHistory.Add oMsg
        If rs!Role = "assistant" Then sLast = CStr(rs!Content)
        rs.MoveNext
    Loop
    rs.Close

    DoCmd.Close acForm, HISTORY_FORM

    ' 更新主窗体
    If Not FormExists(AI_FORM) Then
        MsgBox "请先运行 CreateAIForm 创建 AI 窗体。", vbInformation
        Exit Function
    End If

    DoCmd.OpenForm AI_FORM, acNormal
    Dim frmAI As Form
    Set frmAI = Forms(AI_FORM)
    RebuildChatHtmlFromHistory
    frmAI!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frmAI!txtAnswer.Value = m_sChatHtml
    ScrollAnswerToEnd frmAI
    Dim lTurns As Long, ixH As Long
    lTurns = 0
    For ixH = 1 To m_colHistory.Count
        If m_colHistory(ixH)("role") = "user" Then lTurns = lTurns + 1
    Next ixH
    frmAI!lblMsg.Caption = "已加载历史会话 (第 " & lTurns & " 轮对话)"
End Function

'====================================================
' 删除历史对话
'====================================================
Public Function btnDeleteSession_Click()
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(HISTORY_FORM)

    Dim sSessId As String
    sSessId = Nz(frm!cboSession, "")
    If Len(sSessId) = 0 Then
        MsgBox "请先选择一个会话。", vbInformation
        Exit Function
    End If

    If MsgBox("确定要删除此会话记录吗？", vbQuestion + vbYesNo) = vbNo Then
        Exit Function
    End If

    Dim db  As Object ' DAO.Database
    Set db = CurrentDb
    db.Execute "DELETE FROM " & HISTORY_TABLE & _
               " WHERE SessionID='" & Replace(sSessId, "'", "''") & "'"

    frm!cboSession.Value = Null
    frm!txtHistoryDetail.TextFormat = acTextFormatPlain
    frm!txtHistoryDetail.Value = ""
    RefreshSessionList
End Function

'====================================================
' 按钮入口
'====================================================
Public Sub Askai()
    Dim frm As Form
    Set frm = Screen.ActiveForm

    If Len(Trim$(Nz(frm!txtQ, ""))) = 0 Then
        MsgBox "请输入问题。", vbInformation
        Exit Sub
    End If

    Dim sQuestion As String
    sQuestion = CStr(frm!txtQ)

    ' 获取选择的 AI 提供商
    Dim sProvider As String
    On Error Resume Next
    sProvider = Nz(frm!cboProvider, "DeepSeek Pro")
    On Error GoTo 0

    Dim sUrl As String, sKey As String, sModel As String
    GetProviderConfig sProvider, sUrl, sKey, sModel

    ' 自定义端点校验
    If sProvider = "自定义" Then
        If Len(sUrl) = 0 Or Len(sKey) = 0 Or Len(sModel) = 0 Then
            MsgBox "请填写自定义 API 的 URL、Key 和模型名称。", vbInformation
            Exit Sub
        End If
    End If

    Dim sSystemPrompt As String
    sSystemPrompt = GetSystemPromptFromForm(frm)

    Dim sReasoningEffort As String
    sReasoningEffort = GetReasoningEffortFromForm(frm)

    ' 初始化并添加用户消息到历史
    InitHistory

    ' 每次发送前都从历史重建显示 HTML, 避免模块状态重置导致前文丢失
    RebuildChatHtmlFromHistory

    Dim oUserMsg As Object
    Set oUserMsg = CreateObject("Scripting.Dictionary")
    oUserMsg.Add "role", "user"
    oUserMsg.Add "content", sQuestion
    m_colHistory.Add oUserMsg
    m_sLastAnswer = ""
    m_sStreamingAnswer = ""

    ' 追加用户气泡到对话视图
    RebuildChatHtmlFromHistory
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = m_sChatHtml
    ScrollAnswerToEnd frm

    ' curl.exe 从 Windows 10 1803 开始内置
    If Dir(Environ$("SystemRoot") & "\System32\curl.exe") <> "" Then
        StreamWithCurl frm, sQuestion, sUrl, sKey, sModel, sSystemPrompt, sReasoningEffort
    Else
        SyncWithTypewriter frm, sQuestion, sUrl, sKey, sModel, sSystemPrompt, sReasoningEffort
    End If

    ' 添加助手回复到历史
    If Len(m_sLastAnswer) > 0 Then
        Dim oAsstMsg As Object
        Set oAsstMsg = CreateObject("Scripting.Dictionary")
        oAsstMsg.Add "role", "assistant"
        oAsstMsg.Add "content", m_sLastAnswer
        m_colHistory.Add oAsstMsg
        m_sStreamingAnswer = ""

        ' AI 回复写回历史后重新渲染整段对话, 保证第一轮及后续内容不丢
        RebuildChatHtmlFromHistory
        frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
        frm!txtAnswer.Value = m_sChatHtml
        ScrollAnswerToEnd frm
    End If

    ' 保存到数据库
    If Len(m_sLastAnswer) > 0 Then
        Dim sProviderSave As String
        sProviderSave = Nz(frm!cboProvider, "DeepSeek Pro")
        SaveMessageToDb m_sSessionId, sProviderSave, "user", sQuestion
        SaveMessageToDb m_sSessionId, sProviderSave, "assistant", m_sLastAnswer
    End If

    ' 清空输入框。焦点保留在 txtAnswer, 让鼠标滚轮可以直接滚动回答区。
    frm!txtQ.Value = ""

    ' 更新状态显示对话轮数
    Dim lTurns As Long
    Dim ix As Long
    lTurns = 0
    For ix = 1 To m_colHistory.Count
        If m_colHistory(ix)("role") = "user" Then lTurns = lTurns + 1
    Next ix
    If Len(m_sLastAnswer) > 0 Then
        frm!lblMsg.Caption = "回答完成。 (共 " & Len(m_sLastAnswer) & " 字符, 第 " & lTurns & " 轮对话)"
    End If
End Sub

'====================================================
' 方案A: 真流式 — curl 子进程做 SSE 请求
'
' 原理:
'   1. 将请求体写入临时 JSON 文件
'   2. 用 Shell 启动 curl, 以 SSE 流式接收, 输出到临时文件
'   3. VBA 每 80ms 轮询临时文件, 读取新增内容
'   4. 解析 SSE data 行, 提取 delta.content
'   5. 实时更新文本框 (真正的边接收边显示)
'   6. 收到 [DONE] 后转为 Markdown 富文本
'====================================================
Private Sub StreamWithCurl(frm As Form, ByVal sQuestion As String, _
                           ByVal sUrl As String, ByVal sKey As String, _
                           ByVal sModel As String, ByVal sSystemPrompt As String, _
                           ByVal sReasoningEffort As String)
    On Error GoTo ErrHandler

    ' --- 准备临时文件 ---
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

    ' 构建请求体 (stream=true)
    Dim sBody As String
    sBody = BuildRequestBody(sQuestion, sModel, True, m_colHistory, sSystemPrompt, sReasoningEffort)

    ' 写入请求体文件 (UTF-8 无 BOM)
    WriteUTF8NoBom sTmpBody, sBody

    ' 删除旧响应文件
    On Error Resume Next
    Kill sTmpResp
        Kill sTmpErr
        Kill sTmpDone
    Err.Clear
    On Error GoTo ErrHandler

    ' --- 启动 curl ---
    Dim sCurl As String
        sCurl = """" & Environ$("SystemRoot") & "\System32\curl.exe"" " & _
                "--http1.1 -sS -N --no-buffer " & _
                "-X POST """ & sUrl & """ " & _
                "-H ""Content-Type: application/json; charset=utf-8"" " & _
                "-H ""Authorization: Bearer " & sKey & """ " & _
                "-H ""Accept: text/event-stream"" " & _
                "--data-binary @""" & sTmpBody & """"

        Dim sCmd As String
        sCmd = "cmd /c (" & sCurl & " 1>""" & sTmpResp & """ 2>""" & sTmpErr & """) & echo done>""" & sTmpDone & """"
        Shell sCmd, vbHide

    ' --- UI 初始化 ---
    DoCmd.Hourglass True
    frm!lblMsg.Caption = "AI 正在思考..."
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = m_sChatHtml & BuildAiStreamingBubbleHtml("", True)
    frm.Repaint
    ScrollAnswerToEnd frm

    ' --- 轮询响应文件 ---
    Dim sFullText As String     ' 累积的完整回答
    Dim lLastRawLen As Long     ' 上次读到的原始文本长度
    Dim sngStart As Single      ' 开始时间
    Dim sngLastUI As Single     ' 上次 UI 刷新时间
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
    sCursor = ChrW$(&H258C)    ' ▌

    Do
        DoEvents
        Sleep 80                ' 80ms 一轮

        ' 读取临时文件 (UTF-8)
        sAll = ReadFileAsUTF8(sTmpResp)

        ' 有新内容
        If Len(sAll) > lLastRawLen Then
            lLastRawLen = Len(sAll)

            ' 检查是否结束
            If InStr(sAll, "[DONE]") > 0 Then bDone = True

            ' 重新解析全部 SSE 数据 (简单可靠, 不怕截断)
            Dim sNewFull As String
            sNewFull = ParseSSEChunk(sAll)

            If Len(sNewFull) > Len(sFullText) Then
                sFullText = sNewFull

                ' 首次收到内容
                If Not bFirstToken Then
                    bFirstToken = True
                    DoCmd.Hourglass False
                    frm!lblMsg.Caption = "正在输出..."
                End If

                ' 更新显示 (流式气泡 + 光标)
                m_sStreamingAnswer = sFullText
                frm!txtAnswer.Value = m_sChatHtml & BuildAiStreamingBubbleHtml(sFullText, True)
                frm.Repaint
                ScrollAnswerToEnd frm
                sngLastUI = Timer
            End If
        End If

        bProcDone = (Dir$(sTmpDone) <> "")
        If bDone Then Exit Do
        If bProcDone Then Exit Do

        ' 超时 180 秒
        Dim sngElapsed As Single
        sngElapsed = Timer - sngStart
        If sngElapsed < 0 Then sngElapsed = sngElapsed + 86400  ' 跨午夜
        If sngElapsed > 180 Then
            frm!lblMsg.Caption = "请求超时。"
            sErr = ReadFileAsUTF8(sTmpErr)
            If Len(sErr) > 0 Then
                MsgBox "请求超时。curl 输出:" & vbCrLf & Left$(sErr, 1000), vbExclamation
            Else
                MsgBox "请求超时 (180秒)。", vbExclamation
            End If
            Exit Do
        End If
    Loop

    ' --- 最终显示: Markdown 富文本 ---
    DoCmd.Hourglass False
    m_sLastAnswer = sFullText
    m_sStreamingAnswer = sFullText
    If Len(sFullText) > 0 Then
        ' 将本轮 AI 气泡固化到会话 HTML
        m_sChatHtml = m_sChatHtml & BuildAiBubbleHtml(sFullText)
        frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
        frm!txtAnswer.Value = m_sChatHtml
        frm!lblMsg.Caption = "回答完成。 (共 " & Len(sFullText) & " 字符)"
        ScrollAnswerToEnd frm
    Else
        ' 可能是错误响应: 回退成纯文本显示错误, 不影响会话 HTML
        sAll = ReadFileAsUTF8(sTmpResp)
        sErr = ReadFileAsUTF8(sTmpErr)
        frm!txtAnswer.TextFormat = acTextFormatPlain
        If Len(sErr) > 0 Then
            frm!txtAnswer.Value = "curl 错误:" & vbCrLf & Left$(sErr, 1500)
            frm!lblMsg.Caption = "curl 执行失败。"
        ElseIf Len(sAll) > 0 Then
            frm!txtAnswer.Value = "请求失败:" & vbCrLf & Left$(sAll, 1000)
            frm!lblMsg.Caption = "完成，但返回内容不是有效 SSE。"
        Else
            frm!txtAnswer.Value = "(未收到回答)"
            frm!lblMsg.Caption = "curl 已结束，但未收到内容。"
        End If
    End If
    frm.Repaint

    ' 清理临时文件
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
' 方案B 兜底: 同步请求 + 打字机效果
' (curl 不可用时自动使用)
'====================================================
Private Sub SyncWithTypewriter(frm As Form, ByVal sQuestion As String, _
                               ByVal sUrl As String, ByVal sKey As String, _
                               ByVal sModel As String, ByVal sSystemPrompt As String, _
                               ByVal sReasoningEffort As String)
    On Error GoTo ErrHandler

    Dim sBody As String
    sBody = BuildRequestBody(sQuestion, sModel, False, m_colHistory, sSystemPrompt, sReasoningEffort)

    DoCmd.Hourglass True
    frm!lblMsg.Caption = "AI 正在思考..."
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = m_sChatHtml & BuildAiStreamingBubbleHtml("", True)
    frm.Repaint
    ScrollAnswerToEnd frm

    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts 5000, 10000, 30000, 180000
    xmlHttp.Open "POST", sUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    xmlHttp.setRequestHeader "Authorization", "Bearer " & sKey
    xmlHttp.send sBody

    If xmlHttp.Status <> 200 Then
        frm!lblMsg.Caption = "请求失败: HTTP " & xmlHttp.Status
        MsgBox "HTTP " & xmlHttp.Status & vbCrLf & _
               Left$(xmlHttp.responseText, 500), vbExclamation
        GoTo ExitHere
    End If

    Dim oJson As Object
    Set oJson = JsonConverter.ParseJson(xmlHttp.responseText)
    Dim sAnswer As String
    sAnswer = oJson("choices")(1)("message")("content")
    m_sLastAnswer = sAnswer
    Set xmlHttp = Nothing
    DoCmd.Hourglass False

    If Len(sAnswer) = 0 Then
        MsgBox "API 返回内容为空。", vbExclamation
        GoTo ExitHere
    End If

    frm!lblMsg.Caption = "正在输出..."
    TypewriterShow frm, sAnswer

    ' 固化本轮 AI 气泡
    m_sChatHtml = m_sChatHtml & BuildAiBubbleHtml(sAnswer)
    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = m_sChatHtml
    frm!lblMsg.Caption = "回答完成。 (共 " & Len(sAnswer) & " 字符)"
    ScrollAnswerToEnd frm

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
' 打字机效果 (方案B 使用, 速度自适应)
'====================================================
Private Sub TypewriterShow(frm As Form, ByVal sText As String)
    Dim lTotal As Long
    Dim lPos As Long
    Dim lStep As Long
    Dim lDelay As Long

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

    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText

    For lPos = lStep To lTotal Step lStep
        m_sStreamingAnswer = Left$(sText, lPos)
        frm!txtAnswer.Value = m_sChatHtml & BuildAiStreamingBubbleHtml(Left$(sText, lPos), True)
        frm.Repaint
        ScrollAnswerToEnd frm
        DoEvents
        Sleep lDelay
    Next lPos

    m_sStreamingAnswer = sText
    frm!txtAnswer.Value = m_sChatHtml & BuildAiStreamingBubbleHtml(sText, False)
    frm.Repaint
End Sub

'====================================================
' SSE 解析: 提取所有 data 行的 delta.content
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
' 用 JsonConverter 解析单条 SSE JSON
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
' 统一构建 DeepSeek 请求体
' 使用 JsonConverter 序列化, 避免手工拼 JSON 出错
'====================================================
Private Function BuildRequestBody(ByVal sQuestion As String, _
                                  ByVal sModel As String, _
                                  Optional ByVal bStream As Boolean = False, _
                                  Optional ByVal colHist As Collection = Nothing, _
                                  Optional ByVal sSystemPrompt As String = "", _
                                  Optional ByVal sReasoningEffort As String = "") As String
    Dim oRoot As Object
    Dim colMessages As Collection
    Dim oMsg As Object
    Dim vHistMsg As Variant

    Set oRoot = CreateObject("Scripting.Dictionary")
    Set colMessages = New Collection

    If Len(Trim$(sSystemPrompt)) > 0 Then
        Set oMsg = CreateObject("Scripting.Dictionary")
        oMsg.Add "role", "system"
        oMsg.Add "content", Trim$(sSystemPrompt)
        colMessages.Add oMsg
    End If

    If Not colHist Is Nothing Then
        For Each vHistMsg In colHist
            colMessages.Add vHistMsg
        Next vHistMsg
    Else
        Set oMsg = CreateObject("Scripting.Dictionary")
        oMsg.Add "role", "user"
        oMsg.Add "content", sQuestion
        colMessages.Add oMsg
    End If

    oRoot.Add "model", sModel
    oRoot.Add "messages", colMessages
    oRoot.Add "temperature", 0.7
    oRoot.Add "max_tokens", 8192
    sReasoningEffort = NormalizeReasoningEffort(sReasoningEffort)
    If Len(sReasoningEffort) > 0 Then oRoot.Add "reasoning_effort", sReasoningEffort
    If bStream Then oRoot.Add "stream", True

    BuildRequestBody = JsonConverter.ConvertToJson(oRoot)
End Function

'====================================================
' UTF-8 文件写入 (无 BOM, curl 需要)
'====================================================
Private Sub WriteUTF8NoBom(ByVal sPath As String, ByVal sText As String)
    ' 先用 ADODB.Stream 写 UTF-8 (会带 BOM)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText sText
    stm.SaveToFile sPath, 2
    stm.Close
    Set stm = Nothing

    ' 重新读取二进制, 去掉 3 字节 BOM (EF BB BF)
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

    ' 检查 BOM
    If bAll(0) = &HEF And bAll(1) = &HBB And bAll(2) = &HBF Then
        ' 去掉前 3 字节重写
        Dim bNoBom() As Byte
        ReDim bNoBom(lLen - 4)
        Dim j As Long
        For j = 0 To lLen - 4
            bNoBom(j) = bAll(j + 3)
        Next j
        ' 必须先删除旧文件, 否则 Open For Binary 不截断, 会残留尾部字节
        Kill sPath
        f = FreeFile
        Open sPath For Binary Access Write As #f
        Put #f, 1, bNoBom
        Close #f
    End If
End Sub

'====================================================
' 读取临时文件 (UTF-8 -> VBA 字符串)
' 文件可能正被 curl 写入, 失败时返回空串
'====================================================
Private Function ReadFileAsUTF8(ByVal sPath As String) As String
    On Error Resume Next

    ' 检查文件是否存在
    If Dir(sPath) = "" Then
        ReadFileAsUTF8 = ""
        Exit Function
    End If

    ' 读取原始字节
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

    ' 用 ADODB.Stream 将 UTF-8 字节转为 VBA 字符串
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
'#   第三部分: 窗体自动创建                                  #
'#                                                          #
'############################################################

'====================================================
' 创建 AI 问答窗体 frmAI
' 包含: cboProvider, txtQ, txtAnswer(富文本), lblMsg,
'       btnAsk, btnNewChat, 自定义端点字段
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

    ' ========== 配色常量 (DeepSeek/Gemini 风格) ==========
    Dim cBg As Long          ' 主背景 (纯白)
    Dim cSurface As Long     ' 卡片/输入区
    Dim cBorder As Long      ' 柔和边框
    Dim cText As Long         ' 主要文字
    Dim cSubText As Long      ' 次要文字
    Dim cAccent As Long       ' 强调色 (紫蓝)
    Dim cAccentText As Long   ' 强调色文字
    Dim cToolbar As Long      ' 工具栏背景
    Dim cToolBorder As Long   ' 工具栏下边线
    Dim cBtnHover As Long     ' 按钮悬停

    cBg = RGB(255, 255, 255)
    cSurface = RGB(247, 248, 250)
    cBorder = RGB(228, 231, 236)
    cText = RGB(29, 30, 32)
    cSubText = RGB(134, 142, 153)
    cAccent = RGB(78, 108, 254)
    cAccentText = RGB(255, 255, 255)
    cToolbar = RGB(255, 255, 255)
    cToolBorder = RGB(238, 240, 243)
    cBtnHover = RGB(247, 248, 250)

    ' ========== 窗体主体 ==========
    Set frm = CreateForm
    With frm
        .Caption = "AccessAI"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .Width = 14400
    End With

    frm.Section(acDetail).Height = 11200
    frm.Section(acDetail).BackColor = cBg

    ' ========== 顶栏 (白底 + 底部分隔线, 轻量化) ==========

    ' 顶栏背景
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 0, 0, 14400, 620)
    ctl.BackColor = cToolbar
    ctl.BackStyle = 1
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    ' 顶栏底部分隔线
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 0, 600, 14400, 20)
    ctl.BackColor = cToolBorder
    ctl.BackStyle = 1
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    ' --- 标题: 渐变感图标 + 文字 ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 340, 130, 2000, 360)
    ctl.Caption = ChrW(&H2726) & " AccessAI"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 13
    ctl.FontBold = True
    ctl.ForeColor = cAccent
    ctl.BackStyle = 0

    ' --- cboProvider: 模型下拉框 (胶囊形) ---
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 2600, 130, 2600, 360)
    ctl.Name = "cboProvider"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.RowSourceType = "Value List"
    ctl.RowSource = GetProviderRowSource()
    ctl.DefaultValue = """DeepSeek Pro"""
    ctl.LimitToList = True
    ctl.BackColor = cSurface
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboProvider_AfterUpdate()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 5400, 170, 700, 280)
    ctl.Name = "lblReasoningEffort"
    ctl.Caption = "思考"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 6000, 130, 1400, 360)
    ctl.Name = "cboReasoningEffort"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.RowSourceType = "Value List"
    ctl.RowSource = """默认"";""low"";""medium"";""high"";""xhigh"""
    ctl.DefaultValue = """" & Replace(GetSavedReasoningEffort(), """", """""") & """"
    ctl.LimitToList = True
    ctl.BackColor = cSurface
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboReasoningEffort_AfterUpdate()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 6000, 500, 2100, 180)
    ctl.Name = "lblReasoningCostHint"
    ctl.Caption = "高级别可能增加成本"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 7
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    ' --- btnNewChat: 新对话 ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 7700, 130, 2000, 360)
    ctl.Name = "btnNewChat"
    ctl.Caption = ChrW(&H2795) & " 新对话"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = cText
    ctl.BackColor = cSurface
    ctl.OnClick = "=btnNewChat_Click()"

    ' --- btnHistory: 历史记录 ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 9900, 130, 2200, 360)
    ctl.Name = "btnHistory"
    ctl.Caption = " 历史记录"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = cSubText
    ctl.BackColor = cToolbar
    ctl.OnClick = "=btnHistory_Click()"

    ' ========== 自定义端点字段 (默认隐藏, 浅色卡片) ==========

    ' 自定义区域背景
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 250, 680, 13900, 420)
    ctl.Name = "rectCustomBg"
    ctl.BackColor = cSurface
    ctl.BackStyle = 1
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 720, 500, 300)
    ctl.Name = "lblCustomUrl"
    ctl.Caption = "URL"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 900, 710, 3800, 340)
    ctl.Name = "txtCustomUrl"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.BackColor = cBg
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 4900, 720, 450, 300)
    ctl.Name = "lblCustomKey"
    ctl.Caption = "Key"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 5400, 710, 3200, 340)
    ctl.Name = "txtCustomKey"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.BackColor = cBg
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 8850, 720, 600, 300)
    ctl.Name = "lblCustomModel"
    ctl.Caption = "模型"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0
    ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 9500, 710, 4500, 340)
    ctl.Name = "txtCustomModel"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.BackColor = cBg
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.Visible = False

    ' ========== 系统提示词配置 ==========

    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 250, 1130, 13900, 420)
    ctl.Name = "rectSystemPromptBg"
    ctl.BackColor = cSurface
    ctl.BackStyle = 1
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 1170, 1050, 300)
    ctl.Name = "lblSystemPrompt"
    ctl.Caption = "系统提示词"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 1500, 1160, 12500, 340)
    ctl.Name = "txtSystemPrompt"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.BackColor = cBg
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.DefaultValue = """" & Replace(GetSavedSystemPrompt(), """", """""") & """"
    ctl.AfterUpdate = "=txtSystemPrompt_AfterUpdate()"

    ' ========== 当前数据库表/查询分析 ==========

    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 250, 1580, 13900, 420)
    ctl.Name = "rectDbObjectBg"
    ctl.BackColor = cSurface
    ctl.BackStyle = 1
    ctl.BorderColor = cBorder
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 1620, 900, 300)
    ctl.Name = "lblDbObject"
    ctl.Caption = "数据对象"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 1300, 1610, 6600, 340)
    ctl.Name = "cboDbObject"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.RowSourceType = "Value List"
    ctl.RowSource = GetDbObjectRowSource()
    ctl.LimitToList = True
    ctl.BackColor = cBg
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.OnGotFocus = "=cboDbObject_GotFocus()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 8200, 1610, 2200, 340)
    ctl.Name = "btnAnalyzeData"
    ctl.Caption = "分析数据"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = cAccentText
    ctl.BackColor = cAccent
    ctl.OnClick = "=btnAnalyzeData_Click()"

    ' ========== 核心区域 ==========

    ' --- txtAnswer: 回答区 (大面积白底, 极简边框) ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 500, 2080, 13400, 6360)
    ctl.Name = "txtAnswer"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 11
    ctl.ScrollBars = 2
    ctl.BackColor = cBg
    ctl.BorderStyle = 1
    ctl.BorderColor = cToolBorder
    ctl.SpecialEffect = 0
    ctl.Enabled = True
    ctl.Locked = False
    ctl.TabStop = True
    ctl.EnterKeyBehavior = True

    ' --- lblMsg: 状态标签 ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 500, 8500, 13400, 280)
    ctl.Name = "lblMsg"
    ctl.Caption = "选择模型，输入问题后点击发送"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    ' --- 输入区: 圆角感容器 ---
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 400, 8850, 13600, 1500)
    ctl.BackColor = cSurface
    ctl.BorderColor = cBorder
    ctl.BackStyle = 1
    ctl.SpecialEffect = 0

    ' --- txtQ: 问题输入框 ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 550, 9000, 10800, 1200)
    ctl.Name = "txtQ"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 11
    ctl.ScrollBars = 2
    ctl.EnterKeyBehavior = True
    ctl.BackColor = cSurface
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    ' --- btnAsk: 发送按钮 (品牌色胶囊) ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 11600, 9050, 2200, 1100)
    ctl.Name = "btnAsk"
    ctl.Caption = ChrW(&H27A4) & " 发送"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 11
    ctl.FontBold = True
    ctl.ForeColor = cAccentText
    ctl.BackColor = cAccent
    ctl.OnClick = "=btnAsk_Click()"

    ' 保存窗体
    sTmp = frm.Name
    DoCmd.Close acForm, sTmp, acSaveYes
    Set frm = Nothing

    ' 重新打开设计视图设置 TextFormat
    DoCmd.OpenForm sTmp, acDesign
    Forms(sTmp).Controls("txtAnswer").TextFormat = acTextFormatHTMLRichText
    DoCmd.Close acForm, sTmp, acSaveYes

    ' 重命名
    If sTmp <> AI_FORM Then
        DoCmd.Rename AI_FORM, acForm, sTmp
    End If

    ' 同时创建历史记录表
    EnsureHistoryTable

    MsgBox "窗体 [" & AI_FORM & "] 创建成功!" & vbCrLf & vbCrLf & _
           "打开窗体即可使用 AI 问答。", vbInformation
    Exit Sub

Err_Create:
    MsgBox "CreateAIForm: " & Err.Description, vbExclamation
End Sub

'====================================================
' 方案A: 创建 WebBrowser HTML 对话窗体 frmAIWeb
'====================================================
Public Sub CreateAIWebForm()
    On Error GoTo Err_Create
    Dim frm As Form
    Dim ctl As Control
    Dim sTmp As String

    If FormExists(AI_WEB_FORM) Then
        DoCmd.Close acForm, AI_WEB_FORM, acSaveNo
        DoCmd.DeleteObject acForm, AI_WEB_FORM
    End If

    Set frm = CreateForm
    With frm
        .Caption = "AccessAI - Web 对话模式"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .Width = 14400
        .Section(acDetail).Height = 11200
        .Section(acDetail).BackColor = RGB(255, 255, 255)
    End With

    CreateSharedChatControls frm

    If Not TryCreateWebBrowserControl(frm, ctl) Then
        sTmp = frm.Name
        DoCmd.Close acForm, sTmp, acSaveNo
        DoCmd.DeleteObject acForm, sTmp
        MsgBox "无法自动创建 Microsoft Web Browser ActiveX 控件。" & vbCrLf & vbCrLf & _
               "请确认 Access 已启用 ActiveX 控件，并且系统中已注册 Microsoft Web Browser 控件。", vbExclamation
        Exit Sub
    End If

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 100, 10600, 300, 300)
    ctl.Name = "txtAnswer"
    ctl.Visible = False

    sTmp = frm.Name
    DoCmd.Close acForm, sTmp, acSaveYes
    Set frm = Nothing

    DoCmd.OpenForm sTmp, acDesign
    Forms(sTmp).Controls("txtAnswer").TextFormat = acTextFormatHTMLRichText
    DoCmd.Close acForm, sTmp, acSaveYes

    If sTmp <> AI_WEB_FORM Then DoCmd.Rename AI_WEB_FORM, acForm, sTmp
    EnsureHistoryTable
    MsgBox "窗体 [" & AI_WEB_FORM & "] 创建成功!", vbInformation
    Exit Sub

Err_Create:
    MsgBox "CreateAIWebForm: " & Err.Description, vbExclamation
End Sub

Private Sub CreateSharedChatControls(frm As Form)
    Dim ctl As Control
    Dim cBg As Long, cSurface As Long, cBorder As Long, cText As Long, cSubText As Long, cAccent As Long, cAccentText As Long
    cBg = RGB(255, 255, 255)
    cSurface = RGB(247, 248, 250)
    cBorder = RGB(228, 231, 236)
    cText = RGB(29, 30, 32)
    cSubText = RGB(134, 142, 153)
    cAccent = RGB(78, 108, 254)
    cAccentText = RGB(255, 255, 255)

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 340, 130, 2500, 360)
    ctl.Caption = ChrW(&H2726) & " AccessAI Web"
    ctl.FontName = "Microsoft YaHei": ctl.FontSize = 13: ctl.FontBold = True: ctl.ForeColor = cAccent: ctl.BackStyle = 0

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 2600, 130, 2600, 360)
    ctl.Name = "cboProvider": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 10
    ctl.RowSourceType = "Value List": ctl.RowSource = GetProviderRowSource()
    ctl.DefaultValue = """DeepSeek Pro""": ctl.LimitToList = True: ctl.BackColor = cSurface: ctl.ForeColor = cText: ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboProvider_AfterUpdate()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 5400, 170, 700, 280)
    ctl.Name = "lblReasoningEffort": ctl.Caption = "思考": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 8: ctl.ForeColor = cSubText: ctl.BackStyle = 0
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 6000, 130, 1400, 360)
    ctl.Name = "cboReasoningEffort": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9
    ctl.RowSourceType = "Value List": ctl.RowSource = """默认"";""low"";""medium"";""high"";""xhigh"""
    ctl.DefaultValue = """" & Replace(GetSavedReasoningEffort(), """", """""") & """": ctl.LimitToList = True: ctl.BackColor = cSurface: ctl.ForeColor = cText: ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboReasoningEffort_AfterUpdate()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 6000, 500, 2100, 180)
    ctl.Name = "lblReasoningCostHint": ctl.Caption = "高级别可能增加成本": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 7: ctl.ForeColor = cSubText: ctl.BackStyle = 0

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 7700, 130, 2000, 360)
    ctl.Name = "btnNewChat": ctl.Caption = ChrW(&H2795) & " 新对话": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9: ctl.BackColor = cSurface
    ctl.OnClick = "=btnNewChat_Click()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 9900, 130, 2200, 360)
    ctl.Name = "btnHistory": ctl.Caption = " 历史记录": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9: ctl.ForeColor = cSubText: ctl.BackColor = cBg
    ctl.OnClick = "=btnHistory_Click()"

    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 250, 680, 13900, 420)
    ctl.Name = "rectCustomBg": ctl.BackColor = cSurface: ctl.BackStyle = 1: ctl.BorderColor = cBorder: ctl.BorderStyle = 1: ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 900, 710, 3800, 340)
    ctl.Name = "txtCustomUrl": ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 5400, 710, 3200, 340)
    ctl.Name = "txtCustomKey": ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 9500, 710, 4500, 340)
    ctl.Name = "txtCustomModel": ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 720, 500, 300)
    ctl.Name = "lblCustomUrl": ctl.Caption = "URL": ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 4900, 720, 450, 300)
    ctl.Name = "lblCustomKey": ctl.Caption = "Key": ctl.Visible = False
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 8850, 720, 600, 300)
    ctl.Name = "lblCustomModel": ctl.Caption = "模型": ctl.Visible = False

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 1170, 1050, 300)
    ctl.Name = "lblSystemPrompt": ctl.Caption = "系统提示词": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 8: ctl.ForeColor = cSubText: ctl.BackStyle = 0
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 1500, 1160, 12500, 340)
    ctl.Name = "txtSystemPrompt": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9: ctl.BackColor = cBg: ctl.BorderColor = cBorder: ctl.BorderStyle = 1
    ctl.DefaultValue = """" & Replace(GetSavedSystemPrompt(), """", """""") & """": ctl.AfterUpdate = "=txtSystemPrompt_AfterUpdate()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 400, 1620, 900, 300)
    ctl.Name = "lblDbObject": ctl.Caption = "数据对象": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 8: ctl.ForeColor = cSubText: ctl.BackStyle = 0
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 1300, 1610, 6600, 340)
    ctl.Name = "cboDbObject": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9: ctl.RowSourceType = "Value List": ctl.RowSource = GetDbObjectRowSource(): ctl.LimitToList = True
    ctl.OnGotFocus = "=cboDbObject_GotFocus()"
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 8200, 1610, 2200, 340)
    ctl.Name = "btnAnalyzeData": ctl.Caption = "分析数据": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 9: ctl.ForeColor = cAccentText: ctl.BackColor = cAccent
    ctl.OnClick = "=btnAnalyzeData_Click()"

    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 500, 8500, 13400, 280)
    ctl.Name = "lblMsg": ctl.Caption = "选择模型，输入问题后点击发送": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 8: ctl.ForeColor = cSubText: ctl.BackStyle = 0
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 550, 9000, 10800, 1200)
    ctl.Name = "txtQ": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 11: ctl.ScrollBars = 2: ctl.EnterKeyBehavior = True: ctl.BackColor = cSurface: ctl.BorderStyle = 0
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 11600, 9050, 2200, 1100)
    ctl.Name = "btnAsk": ctl.Caption = ChrW(&H27A4) & " 发送": ctl.FontName = "Microsoft YaHei": ctl.FontSize = 11: ctl.FontBold = True: ctl.ForeColor = cAccentText: ctl.BackColor = cAccent
    ctl.OnClick = "=btnAsk_Click()"
End Sub

'====================================================
' 创建对话历史记录窗体 frmChatHistory
'====================================================
Public Sub CreateHistoryForm()
    On Error GoTo Err_CreateH

    Dim frm As Form
    Dim ctl As Control
    Dim sTmp As String

    If FormExists(HISTORY_FORM) Then
        DoCmd.Close acForm, HISTORY_FORM, acSaveNo
        DoCmd.DeleteObject acForm, HISTORY_FORM
    End If

    Dim cBg As Long, cSurface As Long, cBorder As Long
    Dim cText As Long, cSubText As Long, cAccent As Long
    cBg = RGB(255, 255, 255)
    cSurface = RGB(247, 248, 250)
    cBorder = RGB(228, 231, 236)
    cText = RGB(29, 30, 32)
    cSubText = RGB(134, 142, 153)
    cAccent = RGB(78, 108, 254)

    Set frm = CreateForm
    With frm
        .Caption = "对话历史记录"
        .DefaultView = 0
        .ScrollBars = 0
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .AutoCenter = True
        .Width = 14400
    End With

    frm.Section(acDetail).Height = 10200
    frm.Section(acDetail).BackColor = cBg

    ' --- 顶栏背景 + 分隔线 ---
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 0, 0, 14400, 580)
    ctl.BackColor = cBg
    ctl.BackStyle = 1
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 0, 560, 14400, 20)
    ctl.BackColor = RGB(238, 240, 243)
    ctl.BackStyle = 1
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    ' --- 标题 ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 340, 120, 2200, 340)
    ctl.Caption = " 历史记录"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 12
    ctl.FontBold = True
    ctl.ForeColor = cText
    ctl.BackStyle = 0

    ' --- cboSession: 会话下拉框 ---
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 2800, 110, 5800, 360)
    ctl.Name = "cboSession"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.RowSourceType = "Value List"
    ctl.ColumnCount = 2
    ctl.BoundColumn = 1
    ctl.ColumnWidths = "0"
    ctl.LimitToList = True
    ctl.BackColor = cSurface
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboSession_AfterUpdate()"

    ' --- btnLoadSession ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 8800, 110, 2300, 360)
    ctl.Name = "btnLoadSession"
    ctl.Caption = ChrW(&H21BB) & " 加载对话"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = cAccent
    ctl.OnClick = "=btnLoadSession_Click()"

    ' --- btnDeleteSession ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 11300, 110, 2800, 360)
    ctl.Name = "btnDeleteSession"
    ctl.Caption = ChrW(&H2716) & " 删除记录"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = RGB(198, 40, 40)
    ctl.BackColor = cSurface
    ctl.OnClick = "=btnDeleteSession_Click()"

    ' --- txtHistoryDetail: 对话详情 ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 400, 680, 13600, 9300)
    ctl.Name = "txtHistoryDetail"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.ScrollBars = 2
    ctl.BackColor = cBg
    ctl.BorderStyle = 1
    ctl.BorderColor = RGB(238, 240, 243)
    ctl.SpecialEffect = 0
    ctl.Locked = True
    ctl.TabStop = False
    ctl.EnterKeyBehavior = True

    sTmp = frm.Name
    DoCmd.Close acForm, sTmp, acSaveYes
    Set frm = Nothing

    DoCmd.OpenForm sTmp, acDesign
    Forms(sTmp).Controls("txtHistoryDetail").TextFormat = acTextFormatHTMLRichText
    DoCmd.Close acForm, sTmp, acSaveYes

    If sTmp <> HISTORY_FORM Then
        DoCmd.Rename HISTORY_FORM, acForm, sTmp
    End If
    Exit Sub

Err_CreateH:
    MsgBox "CreateHistoryForm: " & Err.Description, vbExclamation
End Sub

'====================================================
' 创建纯 Markdown 查看窗体 frmMarkdownViewer
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
'#   第四部分: 显示/工具函数                                 #
'#                                                          #
'############################################################

'====================================================
' 弹窗显示 Markdown
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
' 写入任意富文本文本框
'====================================================
Public Sub SetTextBoxMarkdown(txt As TextBox, ByVal sMd As String)
    txt.TextFormat = acTextFormatHTMLRichText
    txt.Value = MarkdownToRichText(sMd)
End Sub

'====================================================
' HTML 转义
'====================================================
Private Function EscHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EscHtml = s
End Function

'====================================================
' JSON 字符串转义
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
' 正则工厂
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
' Markdown 辅助判断函数
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
' UTF-8 文件读取
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
' 判断窗体是否存在
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
