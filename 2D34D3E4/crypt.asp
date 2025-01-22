<%
aesKey="429c1d011549b076ca7ab648666e8436"'  key1
macKey="0b6485591c7d5aacd6c016ce98d7dfe0"'  key2
Set utf8 = CreateObject("System.Text.UTF8Encoding") 
Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform") 
Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform") 
Set mac = CreateObject("System.Security.Cryptography.HMACSHA256") 
Set aes = CreateObject("System.Security.Cryptography.RijndaelManaged") 
Set mem = CreateObject("System.IO.MemoryStream")


'加密方法
Function Encrypt(plaintext, aesKey, macKey)
    ' aes.GenerateIV()
	aes.IV = utf8.GetBytes_4("0909881391141391")
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    Set aesEnc = aes.CreateEncryptor_2((aesKeyBytes), aes.IV)
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    macBytes = ComputeMAC(ConcatBytes(aes.IV, cipherBytes), macKeyBytes)
    Encrypt = B64Encode(macBytes) & ":" & B64Encode(aes.IV) & ":" & _
              B64Encode(cipherBytes)
End Function

'解密方法
Function Decrypt(macIVCiphertext, aesKey, macKey)
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(macKey)
    tokens = Split(macIVCiphertext, ":")
    macBytes = B64Decode(tokens(0))
    ivBytes = B64Decode(tokens(1))
    cipherBytes = B64Decode(tokens(2))
    macActual = ComputeMAC(ConcatBytes(ivBytes, cipherBytes), macKeyBytes)
    If Not EqualBytes(macBytes, macActual) Then
        Err.Raise vbObjectError + 1000, "Decrypt()", "Bad MAC"
    End If
    Set aesDec = aes.CreateDecryptor_2((aesKeyBytes), (ivBytes))
    plainBytes = aesDec.TransformFinalBlock((cipherBytes), 0, LenB(cipherBytes))
    Decrypt = utf8.GetString((plainBytes))
End Function

'取兩個數字中最小的
Function Min(a, b)
    Min = a
    If b < a Then Min = b
End Function

'byte轉base64
Function B64Encode(bytes)
    blockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(bytes) - 1 Step blockSize
        length = Min(blockSize, LenB(bytes) - offset)
        b64Block = b64Enc.TransformFinalBlock((bytes), offset, length)
        result1 = result1 & utf8.GetString((b64Block))
    Next
    B64Encode = result1
End Function

'base64轉byte
Function B64Decode(b64Str)
    bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((bytes), 0, LenB(bytes))
End Function

'連線兩個byte陣列
Function ConcatBytes(a, b)
    mem.SetLength(0)
    mem.Write (a), 0, LenB(a)
    mem.Write (b), 0, LenB(b)
    ConcatBytes = mem.ToArray()
End Function

'比較兩個byte陣列是否相等
Function EqualBytes(a, b)
    EqualBytes = False
    If LenB(a) <> LenB(b) Then Exit Function
    diff = 0
    For i = 1 to LenB(a)
        diff = diff Or (AscB(MidB(a, i, 1)) Xor AscB(MidB(b, i, 1)))
    Next
    EqualBytes = Not diff
End Function

'使用HMAC-SHA-256計算訊息身份驗證程式碼
Function ComputeMAC(msgBytes, keyBytes)
    mac.Key = keyBytes
    ComputeMAC = mac.ComputeHash_2((msgBytes))
End Function
%>