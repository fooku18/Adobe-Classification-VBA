VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSIFICATION_API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** Adobe Credentials **'
Private Const USERNAME As String = "###", _
              SECRET As String = "###", _
              CAMPAIGN_TYPE As String = "###", _
              CREATIVE_CHANNEL As String = "###", _
              CAMPAIGN As String = "###", _
              DOMAIN As String = "###", _
              DEFAULT_PROXY As Boolean = ###, _
			  PROXY As String = "###", _
              REPORT_SUITE As String = "###", _
              EMAIL_NOTIFICATION_ADDRESS As String = "###"

Private Function monteCarlo(ByVal num As Byte, ByVal typ As String) As String
    Const RAND_NUM As String = "0123456789", _
          RAND_CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Byte, _
        s As String: s = ""
    Dim mc As String
    
    If typ = "CHR" Then
        mc = RAND_CHR
    Else
        mc = RAND_NUM
    End If
    For i = 1 To num
        Randomize
        s = s & Mid(mc, Int(Rnd() * Len(mc)) + 1, 1)
    Next i
    monteCarlo = s
End Function

Private Function generateCID(ByVal url As String, ByVal creative_placement As String, ByVal creativ_type As String, ByVal creative As String, ByVal campaign_id As String) As Variant
    Const RAND_CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", _
          RAND_NUM As String = "0123456789"
    Dim URL_enc As String, _
        cid As String
    'Dim cat As New Collection: cat.Add "%3F", "?": cat.Add "%26", "&"
    
    cid = "c_dhlmp_ie_dpcom_" & campaign_id & "_" & monteCarlo(5, "NUM") & "_" & monteCarlo(2, "CHR") & "_" & monteCarlo(3, "NUM")
    If InStr(1, url, "?", vbTextCompare) Then
        url = url & "&cid=" & cid
    Else
        url = url & "?cid=" & cid
    End If
    'URL_enc = Replace(url, "?", cat.Item("?"), 1, -1, vbTextCompare)
    'URL_enc = Replace(URL_enc, "&", cat.Item("&"), 1, -1, vbTextCompare)
    URL_enc = url
    generateCID = Array(cid, URL_enc)
    generateCID = Array(cid, URL_enc)
End Function

Private Function createXHR(ByVal method As String, ByVal load As String, Optional ByVal proxy As Boolean = True) As Object
    Dim xhr As Object, _
        cred As String, _
        oMethod As String, _
        oLoad As String, _
        auth As New WSSE
    
    oMethod = method
    oLoad = load
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    cred = auth.generateAuth(USERNAME, SECRET)
    xhr.Open "POST", "https://api.omniture.com/admin/1.4/rest/?method=" + method, False
    xhr.setRequestHeader "X-WSSE", cred
    xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    If proxy Then xhr.setProxy 2, PROXY
    On Error GoTo xhr_send
    xhr.send (load)
    On Error GoTo 0
    On Error GoTo xhr_empty
    Set createXHR = xhr
    On Error GoTo 0
    Exit Function
xhr_send:
    If err.Number = -2147012889 Then
        Set createXHR = createXHR(oMethod, oLoad, False)
    End If
    err.Clear
    err.Raise "1337", "xhr", "XHR Send failed"
    Exit Function
xhr_empty:
    err.Clear
    err.Raise "1338", "xhr", "XHR Empty"
End Function

Private Function splitClassCols(ByVal template As String) As String()
    Dim a1 As Variant, _
        a2 As Variant
    
    On Error GoTo e
    a1 = split(template, "\r\n")
    a2 = split(a1(3), "\t")
    splitClassCols = a2
    Exit Function
e:
    splitClassCols = Array("Error")
End Function

Private Function getTemplate() As String()
    Dim xhr As Object, _
        load As String
        
    load = "{""element"": ""trackingcode"",""encoding"":""utf8"",""rsid_list"":[""" & REPORT_SUITE & """]}"
    Set xhr = createXHR("Classifications.GetTemplate", load, DEFAULT_PROXY)
    getTemplate = splitClassCols(xhr.responseText)
End Function

Private Function getJobId(ByVal res As String) As Long
    If InStr(1, res, "job_id", vbTextCompare) Then
        a = split(res, ":", -1, vbTextCompare)
        getJobId = CLng(Mid(a(1), 1, Len(a(1)) - 1))
    Else
        getJobId = 0
    End If
End Function

Public Function createImport(ByVal import As String) As Variant
    Dim template() As String, _
        load As String, _
        headers As String, _
        xhr As Object, _
        job_id As Long, _
        job_res As Variant, _
        status As Byte
        
    load = "{""description"":""Import ###CAMPAIGN###"",""element"":""trackingcode"",""email_address"":""" & EMAIL_NOTIFICATION_ADDRESS & """,""export_results"":""false""," & _
            """header"":[""###HEADERS###""],""overwrite_conflicts"":""false"",""rsid_list"":[""" & REPORT_SUITE & """]}"
    template = getTemplate()
    load = Replace(load, "###CAMPAIGN###", CAMPAIGN, 1, -1, vbTextCompare)
    load = Replace(load, "###HEADERS###", Join(template, ""","""), 1, -1, vbTextCompare)
    On Error GoTo xhr_error
    Set xhr = createXHR("Classifications.CreateImport", load, DEFAULT_PROXY)
    On Error GoTo 0
    job_id = getJobId(xhr.responseText)
    Debug.Print job_id
    If job_id = 0 Then
        MsgBox "ERROR: CREATE_IMPORT_ERROR", vbCritical + vbOKOnly, "Error:API:CreateImport"
        createImport = 0
        Exit Function
    Else
        job_res = populateImport(job_id, import)
        If job_res(0) <> 0 Then
            status = commitImport(job_id)
            If status = 1 Then
                createImport = job_res(1)
                Exit Function
            End If
        End If
    End If
    createImport = Array(0)
    Exit Function
xhr_error:
    errorHandler
    MsgBox "ERROR:CREATE_IMPORT:" + err.Description, vbCritical + vbOKOnly, "ERROR"
End Function

Private Function populatePrepare(ByVal import As String) As String()
    Dim ar() As String, _
        ar_ret() As String, _
        i As Long, j As Byte, _
        temp() As String
    
    ar = split(import, "\n", -1, vbTextCompare)
    ReDim ar_ret(UBound(ar) - 1, UBound(split(ar(0), ",", -1, vbTextCompare)))
    For i = 0 To UBound(ar) - 1
        temp = split(ar(i), ",", -1, vbTextCompare)
        For j = 0 To UBound(temp)
            ar_ret(i, j) = temp(j)
        Next j
    Next i
    populatePrepare = ar_ret
End Function

Private Function zeroFill(ByVal id As String) As String
    Const LENGTH As Byte = 5
    Dim l As Integer, _
        i As Byte, _
        ret As String
    
    l = Len(id)
    For i = 1 To LENGTH - l
        ret = ret & "0"
    Next i
    zeroFill = ret & id
End Function

Private Function populateImport(ByVal job_id As Long, ByVal import As String) As Variant
    Dim xhr As Object, _
        load As String, _
        populate() As String, _
        cid() As String, _
        row As String, _
        temp() As Variant, _
        i As Long

    populate = populatePrepare(import)
    populate(i, 2) = zeroFill(populate(i, 2))
    For i = 0 To UBound(populate, 1)
        temp = generateCID(populate(i, 0), populate(i, 3), populate(i, 4), populate(i, 5), populate(i, 2))
        ReDim Preserve cid(i)
        cid(i) = temp(1)
        row = row & "{""row"":[""" & temp(0) & """,""" & CAMPAIGN_TYPE & """,""" & CREATIVE_CHANNEL & """," & _
              """" & populate(i, 2) & """,""" & populate(i, 1) & """,""" & populate(i, 6) & """,""" & populate(i, 7) & """," & _
              """" & populate(i, 3) & """,""" & populate(i, 4) & """,""" & populate(i, 5) & """,""" & DOMAIN & """]},"
    Next i
    row = Mid(row, 1, Len(row) - 1)
    load = "{""job_id"":""###JOBID###"",""page"":""1"",""rows"":[###ROWS###]}"
    load = Replace(load, "###JOBID###", job_id, 1, -1, vbTextCompare)
    load = Replace(load, "###ROWS###", row, 1, -1, vbTextCompare)
    On Error GoTo xhr_error
    Set xhr = createXHR("Classifications.PopulateImport", load, DEFAULT_PROXY)
    On Error GoTo 0
    If xhr.responseText = "true" Then
        populateImport = Array(1, cid)
    Else
        populateImport = Array(0)
    End If
    Exit Function
xhr_error:
    errorHandler
    MsgBox "ERROR:POPULATE_IMPORT:" + err.Description, vbCritical + vbOKOnly, "ERROR"
End Function

Private Function commitImport(ByVal job_id As Long) As Byte
    Dim xhr As Object, _
        load As String, _
        res As String
    
    load = "{""job_id"":""###JOBID###""}"
    load = Replace(load, "###JOBID###", job_id, 1, -1, vbTextCompare)
    On Error GoTo xhr_error
    Set xhr = createXHR("Classifications.CommitImport", load, DEFAULT_PROXY)
    On Error GoTo 0
    res = xhr.responseText
    If res = """Erfolg""" Then
       commitImport = 1
       Exit Function
    End If
    commitImport = 0
    Exit Function
xhr_error:
    errorHandler
    MsgBox "ERROR:GET_STATUS:" + err.Description, vbCritical + vbOKOnly, "ERROR"
End Function

Private Function getStatus(ByVal job_id As Long) As String
    Dim xhr As Object, _
        load As String
    
    load = "{""job_id"":""###JOBID###""}"
    load = Replace(load, "###JOBID###", job_id, 1, -1, vbTextCompare)
    On Error GoTo xhr_error
    Set xhr = createXHR("Classifications.GetStatus", load, DEFAULT_PROXY)
    On Error GoTo 0
    getStatus = xhr.responseText
    Exit Function
xhr_error:
    errorHandler
    MsgBox "ERROR:GET_STATUS:" + err.Description, vbCritical + vbOKOnly, "ERROR"
End Function

Private Function errorHandler()
    
End Function
