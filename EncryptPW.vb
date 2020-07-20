Imports System.Security.Cryptography
Friend NotInheritable Class EncryptPW
    ' Copy From 
    ' https://docs.microsoft.com/en-us/dotnet/articles/visual-basic/programming-guide/language-features/strings/walkthrough-encrypting-And-decrypting-strings

    Private TripleDes As New TripleDESCryptoServiceProvider
    Private _HasError As Boolean

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha256 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha256.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function
    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        HasError = False
    End Sub
    Friend Function EncryptData(ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
        Dim _Result As String = ""

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function
    Friend Function DecryptData(ByVal encryptedtext As String) As String

        If encryptedtext.Length = 0 Or encryptedtext Is Nothing Then
            HasError = True
            Return "Error..."
        End If

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        Try
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()
        Catch ex As Exception
            HasError = True
        End Try

        Dim _Result As String = ""

        If HasError Then
            _Result = ""
        Else
            ' Convert the plaintext stream to a string.
            _Result = System.Text.Encoding.Unicode.GetString(ms.ToArray)

        End If

        Return _Result
    End Function

End Class

Public Module PW
    Dim _HasError As Boolean = False

    Friend Sub Encoding()
        _HasError = False

        Dim plainText As String = InputBox("Enter the plain text:")
        Dim password As String = InputBox("Enter the password:")

        Dim wrapper As New EncryptPW(password)
        Dim cipherText As String = wrapper.EncryptData(plainText)

        MsgBox("The cipher text is: " & cipherText)
        My.Computer.FileSystem.WriteAllText(
            My.Computer.FileSystem.SpecialDirectories.MyDocuments &
            "\cipherText.txt", cipherText, False)
    End Sub
    Friend Sub Decoding()
        _HasError = False
        Dim cipherText As String = My.Computer.FileSystem.ReadAllText(
            My.Computer.FileSystem.SpecialDirectories.MyDocuments &
                "\cipherText.txt")
        Dim password As String = InputBox("Enter the password:")
        Dim wrapper As New EncryptPW(password)

        ' DecryptData throws if the wrong password is used.
        Try
            Dim plainText As String = wrapper.DecryptData(cipherText)
            MsgBox("The plain text is: " & plainText)
        Catch ex As CryptographicException
            MsgBox("The data could not be decrypted with the password.")
        End Try
    End Sub
    '=====================================================================================
    Public Function EncryptPassword(_Password As String, _Wrapper As String) As String
        '_Password is Hash of passwaord
        '_Wrapper of Password Hash

        _HasError = False
        Dim wrapper As New EncryptPW(_Wrapper)
        Dim CipherText As String

        Try
            CipherText = wrapper.EncryptData(_Password)
        Catch ex As Exception
            HasError = True
            MsgBox("Password can not be Encrypt.... Return Nil value")
            CipherText = ""
        End Try


        Return CipherText
    End Function
    Public Function DecryptPassword(_Password As String, _Wrapper As String) As String
        _HasError = False
        Dim PlainText As String = ""
        Dim wrapper As New EncryptPW(_Wrapper)
        Try
            PlainText = wrapper.DecryptData(_Password)
        Catch ex As CryptographicException
            HasError = True
            MsgBox("The data can not be decrypted with the password.")
            PlainText = ""
        End Try
        Return PlainText
    End Function
    '===================================================================================== PROPERTY
    Public Property HasError As Boolean
        Set(value As Boolean)
            _HasError = value
        End Set
        Get
            Return _HasError
        End Get
    End Property

    Public ReadOnly Property GetPassword(_PasswordHash As String, _PWWrapper As String) As String
        Get
            Return DecryptPassword(_PasswordHash, _PWWrapper)
        End Get
    End Property

End Module
