Imports System
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Namespace Utils

    ''' <summary>
    ''' 加密帮助类 - 用于敏感数据（如 API Key）的加密存储
    ''' 使用 Windows 数据保护 API (DPAPI) 进行加密
    ''' 熵值混合应用盐值 + 机器名，防止提取到另一台机器解密
    ''' </summary>
    Public Class EncryptionHelper

        Private Shared ReadOnly _entropy As Byte() = GenerateEntropy()

        ''' <summary>
        ''' 生成机器相关的确定性熵值：应用盐值 + 机器名哈希
        ''' 【安全修复】移除随机盐值组件，确保跨重启稳定性以正确解密已有数据
        ''' DPAPI 本身已提供足够的加密强度，无需额外随机分量
        ''' </summary>
        Private Shared Function GenerateEntropy() As Byte()
            Using sha256 As SHA256 = SHA256.Create()
                ' 确定性熵：仅使用应用标识 + 机器名（保证每次启动生成相同值）
                Dim rawBase As String = "ExcelAIVSTOAddIn|v2.0|" & Environment.MachineName
                Dim hash As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawBase))
                Return hash
            End Using
        End Function

        ''' <summary>
        ''' 加密字符串
        ''' </summary>
        ''' <param name="plainText">明文</param>
        ''' <returns>Base64 编码的密文</returns>
        Public Shared Function Encrypt(plainText As String) As String
            If String.IsNullOrEmpty(plainText) Then
                Return String.Empty
            End If

            Try
                Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
                Dim encryptedBytes As Byte() = ProtectedData.Protect(
                    plainBytes,
                    _entropy,
                    DataProtectionScope.CurrentUser
                )
                Return Convert.ToBase64String(encryptedBytes)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"加密失败: {ex.Message}")
                Throw New InvalidOperationException("加密数据失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 解密字符串
        ''' </summary>
        ''' <param name="cipherText">Base64 编码的密文</param>
        ''' <returns>明文</returns>
        Public Shared Function Decrypt(cipherText As String) As String
            If String.IsNullOrEmpty(cipherText) Then
                Return String.Empty
            End If

            Try
                Dim encryptedBytes As Byte() = Convert.FromBase64String(cipherText)
                Dim plainBytes As Byte() = ProtectedData.Unprotect(
                    encryptedBytes,
                    _entropy,
                    DataProtectionScope.CurrentUser
                )
                Return Encoding.UTF8.GetString(plainBytes)
            Catch ex As CryptographicException
                System.Diagnostics.Debug.WriteLine($"解密失败（可能是数据损坏或权限问题）: {ex.Message}")
                Return String.Empty
            Catch ex As FormatException
                System.Diagnostics.Debug.WriteLine($"解密失败（无效的 Base64 格式）: {ex.Message}")
                Return String.Empty
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"解密失败: {ex.Message}")
                Throw New InvalidOperationException("解密数据失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 检查字符串是否已加密（启发式检查）
        ''' 【改进】增加 DPAPI 密文特征检测：有效 Base64 + 长度>=16 + 非纯 ASCII 可打印文本
        ''' 注意：此方法为快速预检，最终以 Decrypt 是否抛异常为准
        ''' </summary>
        ''' <param name="text">要检查的文本</param>
        ''' <returns>是否可能是加密文本</returns>
        Public Shared Function IsEncrypted(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then
                Return False
            End If

            ' 检查是否是有效的 Base64 字符串
            Try
                Dim bytes As Byte() = Convert.FromBase64String(text)
                ' 加密后的数据通常有一定长度（DPAPI 最小输出约 32 字节 → Base64 后 >= 44 字符）
                If bytes.Length < 32 Then Return False

                ' 【改进】额外检查：DPAPI 加密数据包含非打印字节比例高，
                ' 如果解码后的 bytes 全部是可打印 ASCII 字符，大概率是普通 Base64 编码文本而非密文
                Dim nonPrintableCount As Integer = 0
                For i As Integer = 0 To Math.Min(bytes.Length - 1, 255)
                    Dim b As Integer = bytes(i)
                    If b < 32 OrElse b > 126 Then
                        nonPrintableCount += 1
                    End If
                Next
                ' 前 256 字节中超过 20% 为非可打印字符，更可能是加密数据
                Return nonPrintableCount > CInt(Math.Min(bytes.Length, 256) * 0.2)
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 安全地清除字符串内容（减少内存中明文残留）
        ''' </summary>
        ''' <param name="text">要清除的字符串</param>
        Public Shared Sub SecureClear(ByRef text As String)
            If text IsNot Nothing Then
                ' 注意：这并不能完全保证清除内存，但可以减少风险
                text = New String("0"c, text.Length)
                text = String.Empty
            End If
        End Sub

    End Class

End Namespace
