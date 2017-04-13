VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const chrsz As Byte = 8, _
              b64pad = "="

Private Function shr(ByVal Value As Long, ByVal shift As Byte) As Long
    Dim i As Byte
    shr = Value
    If shift > 0 Then
        shr = Int(shr / (2 ^ shift))
    End If
End Function

Private Function shl(ByVal Value As Long, ByVal shift As Byte) As Long
    shl = Value
    If shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

Private Function str2binb(ByVal s As String) As Long()
    Dim bin() As Long, _
        i As Integer, _
        temp As Byte
    
    ReDim bin(0)
    Mask = shl(1, chrsz) - 1
    For i = 0 To (Len(s) * chrsz) - 1 Step chrsz
        temp = shr(i, 5)
        If temp > UBound(bin) Then
            ReDim Preserve bin(temp)
        End If
        bin(temp) = bin(temp) Or shl(Asc(Mid(s, (i / chrsz) + 1, 1)) And Mask, 24 - i Mod 32)
    Next i
    str2binb = bin
End Function

Private Function rol(ByVal num As Long, ByVal cnt As Byte) As Long
    Dim js As New ScriptControl, _
        ret As Long
    
    js.Language = "JScript"
    js.AddCode "function shift(num,cnt) {return (num << cnt) | (num >>> (32 - cnt))};"
    ret = js.Run("shift", num, cnt)
    rol = ret
End Function

Private Function safe_add(ByVal x As Long, ByVal y As Long) As Long
    Dim lsw As Long, _
        msw As Long
    Const c As Long = 65535
    lsw = (x And c) + (y And c)
    msw = (shr(x, 16)) + (shr(y, 16)) + (shr(lsw, 16))
    safe_add = (shl(msw, 16)) Or (lsw And c)
End Function

Private Function sha1_ft(ByVal t As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    If t < 20 Then
        sha1_ft = (b And c) Or ((Not b) And d)
    ElseIf t < 40 Then
        sha1_ft = b Xor c Xor d
    ElseIf t < 60 Then
        sha1_ft = (b And c) Or (b And d) Or (c And d)
    Else
        sha1_ft = b Xor c Xor d
    End If
End Function

Private Function sha1_kt(ByVal t As Long) As Long
    If t < 20 Then
        sha1_kt = 1518500249
    ElseIf t < 40 Then
        sha1_kt = 1859775393
    ElseIf t < 60 Then
        sha1_kt = -1894007588
    Else
        sha1_kt = -899497514
    End If
End Function

Private Function core_sha1(ByRef x() As Long, ByVal l As Integer) As Long()
    Const CON As Byte = 128
    Dim a As Long: a = 1732584193
    Dim b As Long: b = -271733879
    Dim c As Long: c = -1732584194
    Dim d As Long: d = 271733878
    Dim e As Long: e = -1009589776
    Dim olda As Long, _
        oldb As Long, _
        oldc As Long, _
        oldd As Long, _
        olde As Long, _
        t As Long
    Dim w(79) As Long, _
        i As Integer, _
        j As Integer
    Dim res(4) As Long, _
        al As Byte
    
    al = shr(l, 5)
    ReDim Preserve x(al)
    x(al) = x(al) Or shl(CON, 24 - l Mod 32)
    If UBound(x) < shl(shr(l + 64, 9), 4) + 15 Then ReDim Preserve x(shl(shr(l + 64, 9), 4) + 15)
    x(shl(shr(l + 64, 9), 4) + 15) = l
    
    For i = 0 To UBound(x) - 1 Step 16
        olda = a
        oldb = b
        oldc = c
        oldd = d
        olde = e
        For j = 0 To 79
            If j < 16 Then
                w(j) = x(i + j)
            Else
                w(j) = rol(w(j - 3) Xor w(j - 8) Xor w(j - 14) Xor w(j - 16), 1)
            End If
            t = safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)), safe_add(safe_add(e, w(j)), sha1_kt(j)))
            e = d
            d = c
            c = rol(b, 30)
            b = a
            a = t
        Next j
        a = safe_add(a, olda)
        b = safe_add(b, oldb)
        c = safe_add(c, oldc)
        d = safe_add(d, oldd)
        e = safe_add(e, olde)
    Next i
    res(0) = a: res(1) = b: res(2) = c: res(3) = d: res(4) = e
    core_sha1 = res
End Function

Private Function binb2b64(ByRef binarray() As Long) As String
    Const xtab As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", _
          c As Byte = 63
    Dim xstr As String, _
        i As Long, _
        j As Long, _
        triplet As Long, _
        temp As Long
    For i = 0 To ((UBound(binarray) + 1) * 4) - 1 Step 3
        If shr(i + 2, 2) > UBound(binarray) Then
            temp = shl(shr(binarray(shr(i, 2)), 8 * (3 - i Mod 4)) And 255, 16) Or _
                    shl((shr(binarray(shr(i + 1, 2)), 8 * (3 - (i + 1) Mod 4))) And 255, 8)
        Else
            temp = shl(shr(binarray(shr(i, 2)), 8 * (3 - i Mod 4)) And 255, 16) Or _
                    shl((shr(binarray(shr(i + 1, 2)), 8 * (3 - (i + 1) Mod 4))) And 255, 8) Or _
                    shr(binarray(shr(i + 2, 2)), 8 * (3 - (i + 2) Mod 4)) And 255
        End If
        triplet = temp
        For j = 0 To 3
            If i * 8 + j * 6 > (UBound(binarray) + 1) * 32 Then
                xstr = xstr & b64pad
            Else
                xstr = xstr & Mid(xtab, (shr(triplet, 6 * (3 - j)) And c) + 1, 1)
            End If
        Next j
    Next i
    binb2b64 = xstr
End Function

Public Function Base64_encode(ByVal s As String) As String
    Dim out As String, _
        i As Integer, _
        l As Integer, _
        c1 As Long, _
        c2 As Long, _
        c3 As Long
    Const Cn1 As Byte = 255, _
          EncodeChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", _
          Cn2 As Byte = 3, _
          Cn3 As Byte = 240, _
          Cn4 As Byte = 15, _
          Cn5 As Byte = 192
    l = Len(s)
    i = 0
    out = ""
    Do While i < l
        c1 = Asc(Mid(s, i + 1, 1)) And Cn1
        i = i + 1
        If i = l Then
            out = out & Mid(EncodeChars, shr(c1, 2) + 1, 1)
            out = out & Mid(EncodeChars, shl(c1 And Cn2, 4) + 1, 1)
            out = out & "=="
            Exit Do
        End If
        c2 = Asc(Mid(s, i + 1, 1))
        i = i + 1
        If i = l Then
            out = out & Mid(EncodeChars, (shr(c1, 2)) + 1, 1)
            out = out & Mid(EncodeChars, (shl(c1 And Cn2, 4) Or shr(c2 And Cn3, 4)) + 1, 1)
            out = out & Mid(EncodeChars, (shl(c2 And Cn4, 2)) + 1, 1)
            Exit Do
        End If
        c3 = Asc(Mid(s, i + 1, 1))
        i = i + 1
        out = out & Mid(EncodeChars, (shr(c1, 2)) + 1, 1)
        out = out & Mid(EncodeChars, (shl(c1 And 3, 4) Or shr(c2 And Cn3, 4)) + 1, 1)
        out = out & Mid(EncodeChars, (shl(c2 And 15, 2) Or shr(c3 And Cn5, 6)) + 1, 1)
        out = out & Mid(EncodeChars, (c3 And 63) + 1, 1)
    Loop
    Base64_encode = out
End Function

Public Function b64_sha1(ByVal s As String) As String
    b64_sha1 = binb2b64(core_sha1(str2binb(s), Len(s) * chrsz))
End Function
