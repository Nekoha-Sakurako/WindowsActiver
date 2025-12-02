Option Strict On
Option Explicit On
Option Infer Off
Imports System
Imports System.Text
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Management

Namespace WindowsActiver

    ''' <summary>
    ''' 精简的 SLMgr 功能组件：安装密钥、激活、设置 KMS、查询激活状态。
    ''' 作为 VB.NET 组件使用，调用方负责捕获异常并展示结果。
    ''' </summary>
    <ComVisible(False)>
    Public Class SLMgrComponent
        Implements IDisposable

        Private ReadOnly _computer As String
        Private ReadOnly _username As String
        Private ReadOnly _password As String
        Private _scope As ManagementScope

        Private Const ServiceClass As String = "SoftwareLicensingService"
        Private Const ProductClass As String = "SoftwareLicensingProduct"
        Private Const WindowsAppId As String = "55c92734-d682-4d71-983e-d6ec3f16059f"
        Private Const DefaultKmsPort As Integer = 1688

        ' 用于 Dispose 模式
        Private _disposed As Boolean = False

        Public Sub New(Optional computer As String = ".", Optional username As String = Nothing, Optional password As String = Nothing)
            _computer = If(String.IsNullOrWhiteSpace(computer), ".", computer)
            _username = username
            _password = password
            Connect()
        End Sub

        Private Sub Connect()
            Dim path As String = String.Format(CultureInfo.InvariantCulture, "\\{0}\root\cimv2", _computer)
            Dim options As ConnectionOptions

            If Not String.IsNullOrEmpty(_username) Then
                options = New ConnectionOptions() With {
                    .Username = _username,
                    .Password = _password,
                    .Impersonation = ImpersonationLevel.Impersonate,
                    .Authentication = AuthenticationLevel.PacketPrivacy,
                    .EnablePrivileges = True
                }
            Else
                options = New ConnectionOptions()
            End If

            _scope = New ManagementScope(path, options)
            _scope.Connect()
        End Sub

        ' Helper: query single service management object (first returned)
        Private Function QueryService(selectClause As String) As ManagementObject
            Dim q As New SelectQuery(String.Format("SELECT {0} FROM {1}", selectClause, ServiceClass))
            Dim searcher As New ManagementObjectSearcher(_scope, q)
            Dim results As ManagementObjectCollection = searcher.Get()
            searcher.Dispose()

            For Each mo As ManagementObject In results
                Return mo
            Next
            Return Nothing
        End Function

        ' Helper: query products collection
        Private Function QueryProducts(selectClause As String, Optional whereClause As String = Nothing) As ManagementObjectCollection
            Dim qText As String = If(String.IsNullOrWhiteSpace(whereClause),
                                     String.Format("SELECT {0} FROM {1}", selectClause, ProductClass),
                                     String.Format("SELECT {0} FROM {1} WHERE {2}", selectClause, ProductClass, whereClause))
            Dim q As New ObjectQuery(qText)
            Dim searcher As New ManagementObjectSearcher(_scope, q)
            Dim results As ManagementObjectCollection = searcher.Get()
            searcher.Dispose()
            Return results
        End Function

        ''' <summary>
        ''' 在目标计算机上安装产品密钥。
        ''' 返回操作消息；出错将抛出 ManagementException / Exception。
        ''' </summary>
        Public Function InstallProductKey(productKey As String) As String
            If String.IsNullOrWhiteSpace(productKey) Then
                Throw New ArgumentException("productKey 不能为空。", NameOf(productKey))
            End If

            Dim svc As ManagementObject = QueryService("Version")
            If svc Is Nothing Then
                Throw New InvalidOperationException("未能获取 SoftwareLicensingService。")
            End If

            ' 调用 InstallProductKey 方法（WMI）
            Dim inParams As ManagementBaseObject = svc.GetMethodParameters("InstallProductKey")
            inParams("ProductKey") = productKey
            svc.InvokeMethod("InstallProductKey", inParams, Nothing)

            ' 刷新许可状态（不抛出主异常）
            Try
                svc.InvokeMethod("RefreshLicenseStatus", Nothing, Nothing)
            Catch ex As ManagementException
                ' 忽略刷新失败
            End Try

            Return $"Installed product key: {productKey}"
        End Function

        ''' <summary>
        ''' 激活产品。activationId 可为空（表示默认 Windows 主 SKU）。
        ''' 返回激活结果信息。
        ''' </summary>
        Public Function ActivateProduct(Optional activationId As String = "") As String
            activationId = If(activationId, String.Empty).ToLowerInvariant()
            Dim sb As New StringBuilder()
            Dim svc As ManagementObject = QueryService("Version")
            If svc Is Nothing Then
                Throw New InvalidOperationException("未能获取 SoftwareLicensingService。")
            End If

            Dim selectClause As String = "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus"
            Dim products As ManagementObjectCollection = QueryProducts(selectClause, "PartialProductKey IS NOT NULL")
            Dim foundAny As Boolean = False

            For Each prod As ManagementObject In products
                Dim prodIdObj As Object = prod("ID")
                Dim prodId As String = If(prodIdObj IsNot Nothing, prodIdObj.ToString().ToLowerInvariant(), String.Empty)

                Dim appIdObj As Object = prod("ApplicationId")
                Dim appId As String = If(appIdObj IsNot Nothing, appIdObj.ToString().ToLowerInvariant(), String.Empty)

                Dim isAddonObj As Object = prod("LicenseIsAddon")
                Dim isAddon As Boolean = False
                If isAddonObj IsNot Nothing Then
                    Boolean.TryParse(isAddonObj.ToString(), isAddon)
                End If

                Dim isTarget As Boolean = False
                If String.IsNullOrEmpty(activationId) Then
                    If String.Equals(appId, WindowsAppId, StringComparison.OrdinalIgnoreCase) AndAlso (isAddon = False) Then
                        isTarget = True
                    End If
                Else
                    If prodId = activationId Then
                        isTarget = True
                    End If
                End If

                If isTarget Then
                    foundAny = True
                    sb.AppendLine($"Activating {If(prod("Name"), "<unknown>")} ({If(prod("ID"), "<unknown>")}) ...")
                    Try
                        prod.InvokeMethod("Activate", Nothing)
                        svc.InvokeMethod("RefreshLicenseStatus", Nothing)
                        Dim lsObj As Object = prod("LicenseStatus")
                        Dim ls As Integer = If(lsObj IsNot Nothing, Convert.ToInt32(lsObj), -1)
                        If ls = 1 Then
                            sb.AppendLine("Product activated successfully.")
                        Else
                            sb.AppendLine($"Activation attempted; LicenseStatus={ls}.")
                        End If
                    Catch mex As ManagementException
                        Throw New InvalidOperationException("激活时 WMI 调用失败: " & mex.Message, mex)
                    End Try

                    If Not String.IsNullOrEmpty(activationId) Then
                        Exit For
                    End If
                End If
            Next

            If Not foundAny Then
                Return "Error: product not found."
            End If
            Return sb.ToString().Trim()
        End Function

        ''' <summary>
        ''' 设置 KMS 服务器名和可选端口。kmsNamePort 支持 "[ipv6]" 或 "host:port" 或 "host" 或 ":port"。
        ''' activationId 可选，若为空则针对本机服务级别设置（默认客户端对象）。
        ''' </summary>
        Public Function SetKmsMachineName(kmsNamePort As String, Optional activationId As String = "") As String
            If kmsNamePort Is Nothing Then kmsNamePort = String.Empty
            activationId = If(activationId, String.Empty).ToLowerInvariant()

            ' 解析 [ipv6]:port 或 host:port 或 :port
            Dim kmsName As String = String.Empty
            Dim kmsPortStr As String = String.Empty
            Dim bracketEnd As Integer = kmsNamePort.IndexOf("]"c)
            If kmsNamePort.StartsWith("[") AndAlso bracketEnd > 0 Then
                If kmsNamePort.Length = bracketEnd + 1 Then
                    kmsName = kmsNamePort
                    kmsPortStr = String.Empty
                Else
                    kmsName = kmsNamePort.Substring(0, bracketEnd + 1)
                    kmsPortStr = kmsNamePort.Substring(bracketEnd + 1).TrimStart(":"c)
                End If
            Else
                Dim colonIdx As Integer = kmsNamePort.IndexOf(":"c)
                If colonIdx >= 0 Then
                    kmsName = kmsNamePort.Substring(0, colonIdx)
                    kmsPortStr = kmsNamePort.Substring(colonIdx + 1)
                Else
                    kmsName = kmsNamePort
                End If
            End If

            Dim target As ManagementObject = Nothing

            If String.IsNullOrEmpty(activationId) Then
                target = QueryService("Version, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceLookupDomain")
            Else
                Dim products As ManagementObjectCollection = QueryProducts("ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceLookupDomain", Nothing)
                For Each p As ManagementObject In products
                    Dim idObj As Object = p("ID")
                    Dim id As String = If(idObj IsNot Nothing, idObj.ToString().ToLowerInvariant(), String.Empty)
                    If id = activationId Then
                        target = p
                        Exit For
                    End If
                Next
            End If

            If target Is Nothing Then
                Throw New InvalidOperationException("未找到用于设置 KMS 的目标对象。")
            End If

            Dim prevName As String = If(target("KeyManagementServiceMachine") IsNot Nothing, target("KeyManagementServiceMachine").ToString(), String.Empty)
            Try
                If Not String.IsNullOrEmpty(kmsName) Then
                    Dim inParams As ManagementBaseObject = target.GetMethodParameters("SetKeyManagementServiceMachine")
                    inParams("KeyManagementServiceMachine") = kmsName
                    target.InvokeMethod("SetKeyManagementServiceMachine", inParams, Nothing)
                End If

                If Not String.IsNullOrEmpty(kmsPortStr) Then
                    Dim parsedPort As Integer = DefaultKmsPort
                    If Integer.TryParse(kmsPortStr, parsedPort) Then
                        Dim inP As ManagementBaseObject = target.GetMethodParameters("SetKeyManagementServicePort")
                        inP("KeyManagementServicePort") = parsedPort
                        target.InvokeMethod("SetKeyManagementServicePort", inP, Nothing)
                    Else
                        Throw New ArgumentException("端口解析失败: " & kmsPortStr)
                    End If
                Else
                    Try
                        target.InvokeMethod("ClearKeyManagementServicePort", Nothing, Nothing)
                    Catch
                        ' 部分 WMI 提供者可能不支持 ClearKeyManagementServicePort
                    End Try
                End If
            Catch mex As ManagementException
                If Not String.IsNullOrEmpty(prevName) Then
                    Try
                        Dim rp As ManagementBaseObject = target.GetMethodParameters("SetKeyManagementServiceMachine")
                        rp("KeyManagementServiceMachine") = prevName
                        target.InvokeMethod("SetKeyManagementServiceMachine", rp, Nothing)
                    Catch
                    End Try
                End If
                Throw New InvalidOperationException("设置 KMS 时失败: " & mex.Message, mex)
            End Try

            Return $"KMS 设置成功: {kmsNamePort}"
        End Function

        ''' <summary>
        ''' 返回简明的激活状态字符串（针对所有产品或指定 activationId）。
        ''' </summary>
        Public Function GetActivationStatus(Optional activationId As String = "") As String
            activationId = If(activationId, String.Empty).ToLowerInvariant()
            Dim sb As New StringBuilder()

            Dim selectClause As String = "ID, Name, Description, PartialProductKey, LicenseStatus, GracePeriodRemaining, VLActivationTypeEnabled"
            Dim products As ManagementObjectCollection = QueryProducts(selectClause, "PartialProductKey IS NOT NULL")

            Dim found As Boolean = False
            For Each prod As ManagementObject In products
                Dim idObj As Object = prod("ID")
                Dim id As String = If(idObj IsNot Nothing, idObj.ToString().ToLowerInvariant(), String.Empty)
                Dim show As Boolean = String.IsNullOrEmpty(activationId) OrElse id = activationId

                If show Then
                    found = True
                    sb.AppendLine($"Name: {If(prod("Name"), "<unknown>")}")
                    sb.AppendLine($"Description: {If(prod("Description"), "<unknown>")}")
                    sb.AppendLine($"Activation ID: {If(prod("ID"), "<unknown>")}")
                    sb.AppendLine($"Partial Key: {If(prod("PartialProductKey"), "<none>")}")
                    sb.AppendLine($"LicenseStatus: {If(prod("LicenseStatus") IsNot Nothing, prod("LicenseStatus").ToString(), "<unknown>")}")
                    sb.AppendLine($"GraceMinutesRemaining: {If(prod("GracePeriodRemaining") IsNot Nothing, prod("GracePeriodRemaining").ToString(), "0")}")
                    sb.AppendLine($"VLActivationTypeEnabled: {If(prod("VLActivationTypeEnabled") IsNot Nothing, prod("VLActivationTypeEnabled").ToString(), "0")}")
                    sb.AppendLine(New String("-"c, 40))
                End If
            Next

            If Not found Then
                Return "Error: product key not found."
            End If

            Return sb.ToString().Trim()
        End Function

#Region "IDisposable Support"
        ' Standard Dispose pattern.
        Protected Overridable Sub Dispose(disposing As Boolean)
            If _disposed Then Return

            If disposing Then
                ' dispose managed state (managed objects).
                If _scope IsNot Nothing Then
                    If TypeOf _scope Is IDisposable Then
                        DirectCast(_scope, IDisposable).Dispose()
                    End If
                    _scope = Nothing
                End If
            End If

            ' TODO: free unmanaged resources (if any) here.
            _disposed = True
        End Sub

        ' This code added to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Try
                Dispose(False)
            Finally
                MyBase.Finalize()
            End Try
        End Sub
#End Region
    End Class
End Namespace