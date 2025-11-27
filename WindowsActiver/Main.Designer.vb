<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class WinActFormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        WinActLabel = New Label()
        WinActExec = New Button()
        WinActUtils = New Button()
        WinActKeyList = New ComboBox()
        WinActServerList = New ComboBox()
        WinActKeyTips = New Label()
        WinActServerTips = New Label()
        WinActBar = New ProgressBar()
        SuspendLayout()
        ' 
        ' WinActLabel
        ' 
        WinActLabel.AutoSize = True
        WinActLabel.Font = New Font("Segoe UI", 24F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        WinActLabel.Location = New Point(62, 56)
        WinActLabel.Name = "WinActLabel"
        WinActLabel.Size = New Size(266, 45)
        WinActLabel.TabIndex = 0
        WinActLabel.Text = "WindowsActiver"
        ' 
        ' WinActExec
        ' 
        WinActExec.Location = New Point(62, 391)
        WinActExec.Name = "WinActExec"
        WinActExec.Size = New Size(75, 23)
        WinActExec.TabIndex = 1
        WinActExec.Text = "激活"
        WinActExec.UseVisualStyleBackColor = True
        ' 
        ' WinActUtils
        ' 
        WinActUtils.Location = New Point(253, 391)
        WinActUtils.Name = "WinActUtils"
        WinActUtils.Size = New Size(75, 23)
        WinActUtils.TabIndex = 2
        WinActUtils.Text = "其他功能"
        WinActUtils.UseVisualStyleBackColor = True
        ' 
        ' WinActKeyList
        ' 
        WinActKeyList.FormattingEnabled = True
        WinActKeyList.Location = New Point(52, 162)
        WinActKeyList.Name = "WinActKeyList"
        WinActKeyList.Size = New Size(288, 25)
        WinActKeyList.TabIndex = 3
        WinActKeyList.Text = "（请选择预设产品密钥或输入现有产品密钥）"
        ' 
        ' WinActServerList
        ' 
        WinActServerList.FormattingEnabled = True
        WinActServerList.Location = New Point(52, 255)
        WinActServerList.Name = "WinActServerList"
        WinActServerList.Size = New Size(288, 25)
        WinActServerList.TabIndex = 4
        WinActServerList.Text = "（请选择预设KMS服务器或输入现有服务器地址）"
        ' 
        ' WinActKeyTips
        ' 
        WinActKeyTips.AutoSize = True
        WinActKeyTips.Location = New Point(52, 142)
        WinActKeyTips.Name = "WinActKeyTips"
        WinActKeyTips.Size = New Size(140, 17)
        WinActKeyTips.TabIndex = 5
        WinActKeyTips.Text = "请在此处选择产品密钥："
        ' 
        ' WinActServerTips
        ' 
        WinActServerTips.AutoSize = True
        WinActServerTips.Location = New Point(52, 235)
        WinActServerTips.Name = "WinActServerTips"
        WinActServerTips.Size = New Size(155, 17)
        WinActServerTips.TabIndex = 6
        WinActServerTips.Text = "请在此处选择KMS服务器："
        ' 
        ' WinActBar
        ' 
        WinActBar.Location = New Point(52, 319)
        WinActBar.Name = "WinActBar"
        WinActBar.Size = New Size(288, 23)
        WinActBar.TabIndex = 7
        ' 
        ' WinActFormMain
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(383, 498)
        Controls.Add(WinActBar)
        Controls.Add(WinActServerTips)
        Controls.Add(WinActKeyTips)
        Controls.Add(WinActServerList)
        Controls.Add(WinActKeyList)
        Controls.Add(WinActUtils)
        Controls.Add(WinActExec)
        Controls.Add(WinActLabel)
        Name = "WinActFormMain"
        Text = "WindowsActiver"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents WinActLabel As Label
    Friend WithEvents WinActExec As Button
    Friend WithEvents WinActUtils As Button
    Friend WithEvents WinActKeyList As ComboBox
    Friend WithEvents WinActServerList As ComboBox
    Friend WithEvents WinActKeyTips As Label
    Friend WithEvents WinActServerTips As Label
    Friend WithEvents WinActBar As ProgressBar

End Class
