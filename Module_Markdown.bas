Option Compare Database
Option Explicit
 
'
' 前置条件:
'   1. 导入 JsonConverter 模块 (VBA-JSON by Tim Hall)
'   2. 工具 -> 引用 -> 勾选 "Microsoft Scripting Runtime"
'   3. 修改各 AI 提供商的 API Key
'
' 支持的 AI 模型:
'   DeepSeek / 通义千问 / 文心一言 / Kimi
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
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------- AI 提供商配置 ----------
' 请将下方 Key 替换为你自己的 API Key

' DeepSeek
Private Const DS_KEY   As String = "sk-XXXXXXXXXXXXXXXXXXXX"
Private Const DS_URL   As String = "https://api.deepseek.com/chat/completions"
Private Const DS_MODEL As String = "deepseek-chat"

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

Private Const AI_FORM   As String = "frmAI"
Private Const MD_FORM   As String = "frmMarkdownViewer"
Private Const TXT_MD    As String = "txtMarkdown"

' ---------- 对话历史记录 ----------
Private m_colHistory As Collection
Private m_sLastAnswer As String
Private m_sSessionId As String

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
'#   支持: DeepSeek / 通义千问 / 文心一言 / Kimi             #
'#   方案A: curl 子进程真流式 (Windows 10 1803+)            #
'#   方案B: 同步请求 + 打字机效果 (兜底)                     #
'#                                                          #
'############################################################

'====================================================
' 根据提供商名称返回 API 配置
'====================================================
Private Sub GetProviderConfig(ByVal sProvider As String, _
                              ByRef sUrl As String, _
                              ByRef sKey As String, _
                              ByRef sModel As String)
    Select Case sProvider
        Case "通义千问"
            sUrl = QW_URL: sKey = QW_KEY: sModel = QW_MODEL
        Case "文心一言"
            sUrl = WX_URL: sKey = WX_KEY: sModel = WX_MODEL
        Case "Kimi"
            sUrl = KM_URL: sKey = KM_KEY: sModel = KM_MODEL
        Case "自定义"
            Dim frmC As Form
            Set frmC = Screen.ActiveForm
            sUrl = Nz(frmC!txtCustomUrl, "")
            sKey = Nz(frmC!txtCustomKey, "")
            sModel = Nz(frmC!txtCustomModel, "")
        Case Else  ' DeepSeek (默认)
            sUrl = DS_URL: sKey = DS_KEY: sModel = DS_MODEL
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
End Sub

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
    frm!txtAnswer.TextFormat = acTextFormatPlain
    frm!txtAnswer.Value = ""
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
    If Len(sLast) > 0 Then
        frmAI!txtAnswer.TextFormat = acTextFormatHTMLRichText
        frmAI!txtAnswer.Value = MarkdownToRichText(sLast)
    End If
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
    sProvider = Nz(frm!cboProvider, "DeepSeek")
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

    ' 初始化并添加用户消息到历史
    InitHistory
    Dim oUserMsg As Object
    Set oUserMsg = CreateObject("Scripting.Dictionary")
    oUserMsg.Add "role", "user"
    oUserMsg.Add "content", sQuestion
    m_colHistory.Add oUserMsg
    m_sLastAnswer = ""

    ' curl.exe 从 Windows 10 1803 开始内置
    If Dir(Environ$("SystemRoot") & "\System32\curl.exe") <> "" Then
        StreamWithCurl frm, sQuestion, sUrl, sKey, sModel
    Else
        SyncWithTypewriter frm, sQuestion, sUrl, sKey, sModel
    End If

    ' 添加助手回复到历史
    If Len(m_sLastAnswer) > 0 Then
        Dim oAsstMsg As Object
        Set oAsstMsg = CreateObject("Scripting.Dictionary")
        oAsstMsg.Add "role", "assistant"
        oAsstMsg.Add "content", m_sLastAnswer
        m_colHistory.Add oAsstMsg
    End If

    ' 保存到数据库
    If Len(m_sLastAnswer) > 0 Then
        Dim sProviderSave As String
        sProviderSave = Nz(frm!cboProvider, "DeepSeek")
        SaveMessageToDb m_sSessionId, sProviderSave, "user", sQuestion
        SaveMessageToDb m_sSessionId, sProviderSave, "assistant", m_sLastAnswer
    End If

    ' 清空输入框
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
                           ByVal sModel As String)
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
    sBody = BuildRequestBody(sQuestion, sModel, True, m_colHistory)

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
    frm!txtAnswer.TextFormat = acTextFormatPlain
    frm!txtAnswer.Value = ""
    frm.Repaint

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

                ' 更新显示
                frm!txtAnswer.Value = sFullText & sCursor
                frm.Repaint
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
    If Len(sFullText) > 0 Then
        frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
        frm!txtAnswer.Value = MarkdownToRichText(sFullText)
        frm!lblMsg.Caption = "回答完成。 (共 " & Len(sFullText) & " 字符)"
    Else
        ' 可能是错误响应
        sAll = ReadFileAsUTF8(sTmpResp)
        sErr = ReadFileAsUTF8(sTmpErr)
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
                               ByVal sModel As String)
    On Error GoTo ErrHandler

    Dim sBody As String
    sBody = BuildRequestBody(sQuestion, sModel, False, m_colHistory)

    DoCmd.Hourglass True
    frm!lblMsg.Caption = "AI 正在思考..."
    frm!txtAnswer.TextFormat = acTextFormatPlain
    frm!txtAnswer.Value = ""
    frm.Repaint

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

    frm!txtAnswer.TextFormat = acTextFormatHTMLRichText
    frm!txtAnswer.Value = MarkdownToRichText(sAnswer)
    frm!lblMsg.Caption = "回答完成。 (共 " & Len(sAnswer) & " 字符)"

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
                                  Optional ByVal colHist As Collection = Nothing) As String
    Dim oRoot As Object
    Dim colMessages As Collection

    Set oRoot = CreateObject("Scripting.Dictionary")

    ' 使用对话历史 (如果有)，否则仅发送当前问题
    If Not colHist Is Nothing Then
        Set colMessages = colHist
    Else
        Dim oMsg As Object
        Set colMessages = New Collection
        Set oMsg = CreateObject("Scripting.Dictionary")
        oMsg.Add "role", "user"
        oMsg.Add "content", sQuestion
        colMessages.Add oMsg
    End If

    oRoot.Add "model", sModel
    oRoot.Add "messages", colMessages
    oRoot.Add "temperature", 0.7
    oRoot.Add "max_tokens", 8192
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
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 2800, 130, 2800, 360)
    ctl.Name = "cboProvider"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 10
    ctl.RowSourceType = "Value List"
    ctl.RowSource = """DeepSeek"";""通义千问"";""文心一言"";""Kimi"";""自定义"""
    ctl.DefaultValue = """DeepSeek"""
    ctl.LimitToList = True
    ctl.BackColor = cSurface
    ctl.ForeColor = cText
    ctl.BorderColor = cBorder
    ctl.AfterUpdate = "=cboProvider_AfterUpdate()"

    ' --- btnNewChat: 新对话 ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 6000, 130, 2200, 360)
    ctl.Name = "btnNewChat"
    ctl.Caption = ChrW(&H2795) & " 新对话"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 9
    ctl.ForeColor = cText
    ctl.BackColor = cSurface
    ctl.OnClick = "=btnNewChat_Click()"

    ' --- btnHistory: 历史记录 ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 8400, 130, 2400, 360)
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

    ' ========== 核心区域 ==========

    ' --- txtAnswer: 回答区 (大面积白底, 极简边框) ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 500, 1180, 13400, 7600)
    ctl.Name = "txtAnswer"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 11
    ctl.ScrollBars = 2
    ctl.BackColor = cBg
    ctl.BorderStyle = 1
    ctl.BorderColor = cToolBorder
    ctl.SpecialEffect = 0
    ctl.Locked = True
    ctl.TabStop = False
    ctl.EnterKeyBehavior = True

    ' --- lblMsg: 状态标签 ---
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 500, 8850, 13400, 280)
    ctl.Name = "lblMsg"
    ctl.Caption = "选择模型，输入问题后点击发送"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 8
    ctl.ForeColor = cSubText
    ctl.BackStyle = 0

    ' --- 输入区: 圆角感容器 ---
    Set ctl = CreateControl(frm.Name, acRectangle, acDetail, , , 400, 9200, 13600, 1500)
    ctl.BackColor = cSurface
    ctl.BorderColor = cBorder
    ctl.BackStyle = 1
    ctl.SpecialEffect = 0

    ' --- txtQ: 问题输入框 ---
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 550, 9350, 10800, 1200)
    ctl.Name = "txtQ"
    ctl.FontName = "Microsoft YaHei"
    ctl.FontSize = 11
    ctl.ScrollBars = 2
    ctl.EnterKeyBehavior = True
    ctl.BackColor = cSurface
    ctl.BorderStyle = 0
    ctl.SpecialEffect = 0

    ' --- btnAsk: 发送按钮 (品牌色胶囊) ---
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 11600, 9400, 2200, 1100)
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