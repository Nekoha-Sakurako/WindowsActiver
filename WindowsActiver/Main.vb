Public Class WinActFormMain
    Dim KeyValue As String
    Dim ServerValue As String
    Dim KeyUninst
    Dim objShell
    Dim FailMsg
    Dim ExecMsg
    Sub Activator()
        objShell = CreateObject("WScript.Shell")
        WinActBar.Value = 0
        objShell.Run("slmgr.vbs /upk", 0, True)
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /ipk " + KeyValue, 0, True)
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /skms " + KeyValue, 0, True)
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /ato", 0, True)
        WinActBar.Value = WinActBar.Value + 25
        'Retry Test
        FailMsg = MsgBox("该Windows副本激活失败，原因未知。", vbCritical + vbRetryCancel, "激活失败")
        If FailMsg = vbRetry Then
            Activator()
        Else
            WinActBar.Value = 0
            WinActExec.Enabled = True
            WinActKeyList.Enabled = True
            WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
            WinActServerList.Enabled = True
            WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
            WinActExec.Text = "点击激活"
            Me.UseWaitCursor = False
        End If
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
        WinActExec.Enabled = True
        WinActKeyList.Enabled = True
        WinActKeyTips.Text = WinActKeyTips.Text.Substring(WinActKeyTips.Text.Length - 11)
        WinActServerList.Enabled = True
        WinActServerTips.Text = WinActServerTips.Text.Substring(WinActServerTips.Text.Length - 13)
        WinActExec.Text = "点击激活"
    End Sub

    Private Sub WinActUtilities_SelectedIndexChanged(sender As Object, e As EventArgs) Handles WinActUtilities.SelectedIndexChanged
        If WinActUtilities.SelectedItem = "备用激活方案…" Then
            KeyValue = ""
            ServerValue = ""
            KeyValue = InputBox("请输入产品密钥：", "备用激活方案")
            ServerValue = InputBox("请输入KMS服务器地址："， "备用激活方案")
            If KeyValue = "" Then
                MsgBox("你尚未输入Windows产品密钥，无法激活。", 48, "激活失败")
            ElseIf ServerValue = "" Then
                MsgBox("你尚未设置KMS服务器，无法激活。", 48, "激活失败")
            Else
                Activator()
            End If
        ElseIf WinActUtilities.SelectedItem = "卸载产品密钥…" Then
            KeyUninst = MsgBox("确认要卸载已安装的产品密钥吗？", vbQuestion + vbYesNo, "二次确认")
            If KeyUninst = vbYes Then
                Me.UseWaitCursor = True
                Me.Text = （"(正在卸载产品密钥……)" & Me.Text）
                Dim objShell
                objShell = CreateObject("WScript.Shell")
                objShell.Run("slmgr.vbs /upk", 0, True)
                Me.UseWaitCursor = False
                Me.Text = Me.Text.Substring(Me.Text.Length - 14)
            End If
        ElseIf WinActUtilities.SelectedItem = "调试功能…" Then
            MsgBox("目前还没有")
        End If
    End Sub
End Class
