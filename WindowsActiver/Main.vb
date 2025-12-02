Imports WindowsActiver
Imports System.Threading.Tasks
Public Class WinActFormMain
    Dim KeyValue As String
    Dim ServerValue As String
    Dim KeyUninst
    Dim ServerUninst
    Dim Uninst
    Dim objShell
    Dim FailMsg
    Dim ExecMsg
    Dim SLProm

    Sub Activator()
        ' 使用模块中的 SLMgrComponent 替代外部 slmgr.vbs 调用
        WinActBar.Value = 0

        Try
            Me.UseWaitCursor = True

            Using comp As New WindowsActiver.SLMgrComponent()
                ' 如果需要保留卸载已安装产品密钥的步骤，可在 SLMgrComponent 中实现对应方法并在此调用
                ' 这里按现有组件能力执行：安装密钥 -> 设置 KMS -> 激活
                WinActBar.Value = WinActBar.Value + 25

                ' 安装产品密钥
                Dim installResult As String = comp.InstallProductKey(KeyValue)
                MsgBox（"产品密钥已安装" & KeyValue）
                WinActBar.Value = WinActBar.Value + 25

                ' 设置 KMS（注意使用 ServerValue）
                MsgBox("服务器安装" & ServerValue)
                Dim kmsResult As String = comp.SetKmsMachineName(ServerValue)
                WinActBar.Value = WinActBar.Value + 25

                ' 激活（返回字符串表示结果）
                Dim actResult As String = comp.ActivateProduct()
                WinActBar.Value = WinActBar.Value + 25

                ' 根据返回结果判断是否成功（组件会返回包含 "activated" 或 "Error" 的文本）
                Dim lower = If(actResult, String.Empty).ToLowerInvariant()
                If lower.Contains("activated") Or lower.Contains("success") Then
                    MsgBox("恭喜，已成功激活该Windows副本。" & vbCrLf & actResult, vbInformation, "激活完毕")
                Else
                    ' 激活未明确成功，询问是否重试（保留原有重试 UX）
                    FailMsg = MsgBox("该Windows副本激活失败，返回信息：" & vbCrLf & actResult & vbCrLf & vbCrLf & "是否重试？", vbCritical + vbRetryCancel, "激活失败")
                    If FailMsg = vbRetry Then
                        ' 重置进度及控件状态后重试
                        WinActBar.Value = 0
                        Activator()
                        Return
                    Else
                        ' 回滚 UI 状态到可用
                        WinActBar.Value = 0
                        WinActExec.Enabled = True
                        WinActKeyList.Enabled = True
                        If WinActKeyTips.Text.Length >= 11 Then
                            WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
                        End If
                        WinActServerList.Enabled = True
                        If WinActServerTips.Text.Length >= 13 Then
                            WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
                        End If
                        WinActExec.Text = "点击激活"
                        Me.UseWaitCursor = False
                        Return
                    End If
                End If
            End Using

            ' 完成后恢复 UI
            WinActBar.Value = 0
            WinActExec.Enabled = True
            WinActKeyList.Enabled = True
            If WinActKeyTips.Text.Length >= 11 Then
                WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
            End If
            WinActServerList.Enabled = True
            If WinActServerTips.Text.Length >= 13 Then
                WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
            End If
            WinActExec.Text = "点击激活"
            Me.UseWaitCursor = False

        Catch ex As Exception
            ' 捕获异常并提示，支持重试
            FailMsg = MsgBox("该Windows副本激活失败，原因：" & vbCrLf & ex.Message, vbCritical + vbRetryCancel, "激活失败")
            If FailMsg = vbRetry Then
                Activator()
            Else
                WinActBar.Value = 0
                WinActExec.Enabled = True
                WinActKeyList.Enabled = True
                If WinActKeyTips.Text.Length >= 11 Then
                    WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
                End If
                WinActServerList.Enabled = True
                If WinActServerTips.Text.Length >= 13 Then
                    WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
                End If
                WinActExec.Text = "点击激活"
                Me.UseWaitCursor = False
            End If
        End Try
    End Sub

    Private Sub WinActExec_Click(sender As Object, e As EventArgs) Handles WinActExec.Click
        If WinActKeyList.Text = "（请选择预设产品密钥或输入现有产品密钥）" Then
            MsgBox("你尚未输入Windows产品密钥，无法激活。", 48, "激活失败")
        ElseIf WinActKeyList.Text = "" Then
            MsgBox("你尚未输入Windows产品密钥，无法激活。", 48, "激活失败")
        ElseIf WinActServerList.Text = "（请选择预设KMS服务器或输入现有服务器地址）" Then
            MsgBox("你尚未设置KMS服务器，无法激活。", 48, "激活失败")
        ElseIf WinActServerList.Text = "" Then
            MsgBox("你尚未设置KMS服务器，无法激活。", 48, "激活失败")
        Else
            ExecMsg = MsgBox("确认要激活Windows副本吗？" & vbCrLf & "你选择的产品密钥是：" + WinActKeyList.Text & vbCrLf & "你选择的KMS服务器是：" + WinActServerList.Text, vbQuestion + vbYesNo, "二次确认")
        End If
        If ExecMsg = vbYes Then
            KeyValue = WinActKeyList.Text
            ServerValue = WinActServerList.Text
            Me.UseWaitCursor = True
            WinActExec.Enabled = False
            WinActKeyList.Enabled = False
            WinActKeyTips.Text = "（激活过程中不可修改） " + WinActKeyTips.Text
            WinActServerList.Enabled = False
            WinActServerTips.Text = "（激活过程中不可修改） " + WinActServerTips.Text
            WinActExec.Text = "正在激活"
            Activator()
        End If
        'if the activator does not execute the reset operate then do following commands
        WinActExec.Enabled = True
        WinActKeyList.Enabled = True
        If WinActKeyTips.Text.Length >= 11 Then
            WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
        End If
        WinActServerList.Enabled = True
        If WinActServerTips.Text.Length >= 13 Then
            WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
        End If
        WinActExec.Text = "点击激活"
    End Sub

    ' 在 UI 线程显示长文本的模态窗体
    Private Sub ShowLargeText(title As String, text As String)
        Dim f As New Form() With {
            .Text = title,
            .Size = New Drawing.Size(800, 600),
            .StartPosition = FormStartPosition.CenterParent
        }
        Dim tb As New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .Dock = DockStyle.Fill,
            .Font = New Drawing.Font("Consolas", 10),
            .WordWrap = False,
            .BackColor = Drawing.SystemColors.Window
        }
        Dim btn As New Button() With {
            .Text = "关闭",
            .Dock = DockStyle.Bottom,
            .Height = 32
        }
        AddHandler btn.Click, Sub(s, e) f.Close()
        tb.Text = If(text, String.Empty)
        f.Controls.Add(tb)
        f.Controls.Add(btn)
        f.ShowDialog(Me)
    End Sub

    ' 异步查询激活状态并在完成后展示
    Private Sub ShowActivationStatus()
        Me.UseWaitCursor = True
        WinActUtilities.Enabled = False

        Task.Run(Function()
                     Dim result As String
                     Try
                         Using comp As New WindowsActiver.SLMgrComponent()
                             result = comp.GetActivationStatus() ' 调用组件（WMI 查询）
                         End Using
                     Catch ex As Exception
                         result = "查询激活状态失败: " & ex.Message
                     End Try

                     ' 回到 UI 线程显示结果
                     Me.Invoke(Sub()
                                   Me.UseWaitCursor = False
                                   WinActUtilities.Enabled = True
                                   ShowLargeText("激活状态", result)
                               End Sub)
                     Return 0
                 End Function)
    End Sub

    Private Sub WinActUtilities_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WinActUtilities.SelectedIndexChanged
        objShell = CreateObject("WScript.Shell")
        If WinActUtilities.SelectedItem = "备用激活方案…" Then
            KeyValue = ""
            ServerValue = ""
            KeyValue = InputBox("请输入产品密钥：", "备用激活方案")
            ServerValue = InputBox("请输入KMS服务器地址：", "备用激活方案")
            If KeyValue = "" Then
                MsgBox("你尚未输入Windows产品密钥，无法激活。", 48, "激活失败")
            ElseIf ServerValue = "" Then
                MsgBox("你尚未设置KMS服务器，无法激活。", 48, "激活失败")
            Else
                ExecMsg = MsgBox("确认要激活Windows副本吗？" & vbCrLf & "你选择的产品密钥是：" + KeyValue & vbCrLf & "你选择的KMS服务器是：" + ServerValue, vbQuestion + vbYesNo, "二次确认")
                If ExecMsg = vbYes Then
                    Activator()
                End If
            End If
        ElseIf WinActUtilities.SelectedItem = "卸载产品密钥…" Then
            KeyUninst = MsgBox("确认要卸载已安装的产品密钥吗？", vbQuestion + vbYesNo, "二次确认")
            If KeyUninst = vbYes Then
                Me.UseWaitCursor = True
                Me.Text = "(正在卸载产品密钥……)" & Me.Text
                objShell.Run("slmgr.vbs /upk", 0, True)
                Me.UseWaitCursor = False
                Me.Text = Me.Text.Substring(Me.Text.Length - 14)
            End If
        ElseIf WinActUtilities.SelectedItem = "取消已设置的KMS服务器…" Then
            ServerUninst = MsgBox("确认要取消已设置的KMS服务器吗？", vbQuestion + vbYesNo, "二次确认")
            If ServerUninst = vbYes Then
                Me.UseWaitCursor = True
                Me.Text = "(正在取消设置KMS服务器……)" & Me.Text
                objShell.Run("slmgr.vbs /ckms", 0, True)
                Me.UseWaitCursor = False
                Me.Text = Me.Text.Substring(Me.Text.Length - 14)
            End If
        ElseIf WinActUtilities.SelectedItem = "显示激活状态…" Then
            ' 不再直接弹出 slmgr 窗口，而是使用 SLMgrComponent 查询并在窗体内显示
            ShowActivationStatus()
        ElseIf WinActUtilities.SelectedItem = "查询到期时间…" Then
            objShell.Run("slmgr.vbs /xpr", 0, True)
        ElseIf WinActUtilities.SelectedItem = "取消激活（危险功能！！！）…" Then
            Uninst = MsgBox("确认要取消激活吗？" & vbCrLf & "请三思后再进行选择！", vbQuestion + vbYesNo, "二次确认")
            If Uninst = vbYes Then
                Uninst = MsgBox("你真的确认要取消激活吗？" & vbCrLf & "这是最后一次警告！该操作不可逆！", vbExclamation + vbYesNo, "三次确认")
                If Uninst = vbYes Then
                    Me.UseWaitCursor = True
                    Me.Text = "(正在取消激活……)" & Me.Text
                    objShell.Run("slmgr.vbs /rearm", 0, True)
                    Me.UseWaitCursor = False
                    Me.Text = Me.Text.Substring(Me.Text.Length - 14)
                End If
            End If
        ElseIf WinActUtilities.SelectedItem = "高级功能…" Then
            SLProm = InputBox("请输入指令（内容留空以获取帮助）：")
            objShell.Run("slmgr.vbs /" & SLProm, 0, True)
        ElseIf WinActUtilities.SelectedItem = "调试功能…" Then
            MsgBox("目前还没有")
            MsgBox("恭喜，已成功激活该Windows副本。", 64, "激活完毕")
        End If
    End Sub
End Class
