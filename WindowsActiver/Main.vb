Public Class WinActFormMain
    Sub Activator()
        Dim objShell
        Dim FailMsg
        objShell = CreateObject("WScript.Shell")
        WinActBar.Value = 0
        objShell.Run("slmgr.vbs /upk")
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /ipk " + WinActKeyList.Text, )
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /skms " + WinActServerList.Text, )
        WinActBar.Value = WinActBar.Value + 25
        objShell.Run("slmgr.vbs /ato")
        WinActBar.Value = WinActBar.Value + 25
        'Retry Test
        FailMsg = MsgBox("该Windows副本激活失败，原因未知。", vbCritical + vbRetryCancel, "激活失败")
        If FailMsg = vbRetry Then
            Activator()
        Else
            WinActBar.Value = 0
            WinActExec.Enabled = True
            WinActKeyList.Enabled = True
            'WinActKeyTips.Text = Right(WinActKeyTips.Text, 11)
            WinActServerList.Enabled = True
            'WinActServerTips.Text = Right(WinActServerTips.Text, 13)
            WinActExec.Text = "点击激活"
        End If
    End Sub
    Private Sub WinActExec_Click(sender As Object, e As EventArgs) Handles WinActExec.Click
        Dim ExecMsg
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
            WinActExec.Enabled = False
            WinActKeyList.Enabled = False
            'WinActKeyTips.Text = "（激活过程中不可修改） " + WinActKeyTips.Text
            WinActServerList.Enabled = False
            'WinActServerTips.Text = "（激活过程中不可修改） " + WinActServerTips.Text
            WinActExec.Text = "正在激活"
            Activator()
        End If
        WinActExec.Enabled = True
        WinActKeyList.Enabled = True
        '        WinActKeyTips.Text = Right(WinActKeyTips.Text, 11)
        WinActServerList.Enabled = True
        '        WinActServerTips.Text = Right(WinActServerTips.Text, 13)
        WinActExec.Text = "点击激活"
    End Sub
End Class
