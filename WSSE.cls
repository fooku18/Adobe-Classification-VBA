VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WSSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_u As String, _
        p_s As String, _
        p_n As String, _
        p_c As String

Private Function generateNonce() As String
    Const LENGTH As Integer = 24, _
        chars As String = "0123456789abcdef"
    Dim nonce As String: nonce = ""
    Dim i As Integer
    For i = 1 To LENGTH Step 1
        Randomize
        nonce = nonce & Mid(chars, Int(Rnd * Len(chars)) + 1, 1)
    Next i
    generateNonce = nonce
    'generateNonce = "0123456789abcdef01234567"
End Function

Private Function zeroFill(ByVal s As String) As String
    If Len(s) = 1 Then
        zeroFill = "0" & s
    Else
        zeroFill = s
    End If
End Function

Private Function generateCreated() As String
    Dim dt As Date, _
        y As String, m As String, d As String, h As String, i As String, s As String
    Dim l2utc As New LocalToUTC
    
    dt = l2utc.ConvertLocalToGMT(Now)
    y = Year(dt)
    m = Month(dt)
    d = Day(dt)
    h = Hour(dt)
    i = Minute(dt)
    s = Second(dt)
    
    'generateCreated = "2017-03-30T21:30:15Z"
    generateCreated = y & "-" & zeroFill(m) & "-" & zeroFill(d) & "T" & zeroFill(h) & ":" & zeroFill(i) & ":" & zeroFill(s) & "Z"
End Function

Private Function encode() As Collection
    Dim p As New Collection
    Dim sha As New sha, _
        t As String
    
    p.Add p_u, "u"
    p.Add p_s, "s"
    p.Add p_n, "n"
    p.Add p_c, "c"
    p.Add sha.b64_sha1(p.Item("n") & p.Item("c") & p.Item("s")), "d"
    t = p.Item("n")
    p.Remove ("n")
    p.Add sha.Base64_encode(t), "n"
    Set encode = p
End Function

Private Function generateRESTHeaders() As String
    Dim rval As String: rval = ""
    Dim p As New Collection
    
    Set p = encode()
    rval = "UsernameToken"
    rval = rval & " Username=""" & p.Item("u") & ""","
    rval = rval & " PasswordDigest=""" & p.Item("d") & ""","
    rval = rval & " Nonce=""" & p.Item("n") & ""","
    rval = rval & " Created=""" & p.Item("c") & """"
    
    generateRESTHeaders = rval
End Function

Public Function generateAuth(ByVal USERNAME As String, ByVal SECRET As String) As String
    Dim nonce As String, _
        created As String, _
        headers As String
    nonce = generateNonce()
    created = generateCreated()
    p_u = USERNAME
    p_s = SECRET
    p_n = nonce
    p_c = created
    
    generateAuth = generateRESTHeaders()
End Function
