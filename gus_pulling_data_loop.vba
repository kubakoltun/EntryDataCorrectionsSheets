Option Explicit
#If VBA7 Then
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
          ByVal CodePage As Long, _
          ByVal dwFlags As Long, _
          ByVal lpWideCharStr As LongPtr, _
          ByVal cchWideChar As Long, _
          ByVal lpMultiByteStr As String, _
          ByVal cchMultiByte As Long, _
          ByVal lpDefaultChar As LongPtr, _
          ByVal lpUsedDefaultChar As LongPtr) As Long
  Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
          ByVal CodePage As Long, _
          ByVal dwFlags As Long, _
          ByVal lpMultiByteStr As String, _
          ByVal cchMultiByte As Long, _
          ByVal lpWideCharStr As LongPtr, _
          ByVal cchWideChar As Long) As Long
  Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
#Else
  Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
  Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
          ByVal CodePage As Long, _
          ByVal dwFlags As Long, _
          ByVal lpWideCharStr As Long, _
          ByVal cchWideChar As Long, _
          ByVal lpMultiByteStr As String, _
          ByVal cchMultiByte As Long, _
          ByVal lpDefaultChar As Long, _
          ByVal lpUsedDefaultChar As Long) As Long
  Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
          ByVal CodePage As Long, _
          ByVal dwFlags As Long, _
          ByVal lpMultiByteStr As String, _
          ByVal cchMultiByte As Long, _
          ByVal lpWideCharStr As Long, _
          ByVal cchWideChar As Long) As Long
  Private Declare Function GetACP Lib "kernel32" () As Long
#End If


Sub test()
'
' test Makro
Dim api As String
Dim id As String
Dim key As String
Dim res As String
Dim NIP As String
Dim lastRow As Long
Dim currentRow As Integer
Dim i As Variant

api = "test"
key = "abcde1234"
currentRow = 4

lastRow = Cells(Rows.Count, "A").End(xlUp).Row
For Each i In Range("A4:A" & lastRow).Cells
    Sleep 334
    NIP = i.Value
    'leave only numbers
    NIP = Replace(NIP, "-", "")
    NIP = Replace(NIP, " ", "")
    
    'nip len check
    If Len(NIP) > 10 Then
        Range("B" & currentRow).Value = ("Wprowadzony NIP jest za długi, podana ilość znaków - " & Len(NIP))
        currentRow = currentRow + 1
        GoTo NextIteration
    End If
    If Len(NIP) < 10 Then
        Range("B" & currentRow).Value = ("Wprowadzony NIP jest za krótki, podana ilość znaków - " & Len(NIP))
        currentRow = currentRow + 1
        GoTo NextIteration
    End If
    
    'log in
    With CreateObject("winhttp.winhttprequest.5.1")
            .Open "POST", api, False
            .setRequestHeader "Content-Type", "application/soap+xml;charset=UTF-8;"
            .send "" & _
                    "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns=""http://CIS/BIR/PUBL/2014/07"">" & _
                    "<soap:Header xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" & _
                    "<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/Zaloguj</wsa:Action>" & _
                    "<wsa:To>" + api + "</wsa:To>" & _
                    "</soap:Header>" & _
                    "<soap:Body>" & _
                    "<ns:Zaloguj>" & _
                    "<ns:pKluczUzytkownika>" + key + "</ns:pKluczUzytkownika>" & _
                    "</ns:Zaloguj>" & _
                    "</soap:Body>" & _
                    "</soap:Envelope>"
                    
    res = .responseText
     If Len(res) = 0 Then
        Range("B" & currentRow).Value = "Nie można uzyskać sesji z usługi sieciowej GUSu!"
        currentRow = currentRow + 1
        GoTo NextIteration
     End If
    id = Split(res, "ZalogujResult>")(1)
    id = Left(id, Len(id) - 2)
    'Range("C8").Value = id
    End With
    
    'nip pull data
    With CreateObject("winhttp.winhttprequest.5.1")
            .Open "POST", api, False
            .setRequestHeader "Content-Type", "application/soap+xml; charset=UTF-8;"
            .setRequestHeader "sid", id
            .send "" & _
                    "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns=""http://CIS/BIR/PUBL/2014/07"" xmlns:dat=""http://CIS/BIR/PUBL/2014/07/DataContract"">" & _
                        "<soap:Header xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" & _
                            "<wsa:To>" + api + "</wsa:To>" & _
                            "<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/DaneSzukajPodmioty</wsa:Action>" & _
                        "</soap:Header>" & _
                    "<soap:Body>" & _
                        "<ns:DaneSzukajPodmioty>" & _
                            "<ns:pParametryWyszukiwania>" & _
                                "<dat:Nip>" + NIP + "</dat:Nip>" & _
                            "</ns:pParametryWyszukiwania>" & _
                        "</ns:DaneSzukajPodmioty>" & _
                    "</soap:Body>" & _
                    "</soap:Envelope>"
    'Range("B7").Value = .responseText
    res = VBA.Strings.StrConv(.responseBody, vbUnicode)
    'Range("D7").Value = res
    res = tekstCodePageToCodePage(res, 65001, 1250)
    
    res = Replace(res, "&lt;", "<")
    res = Replace(res, "&gt;", ">")
    
    Dim a As String
    
    'errors
    If (InStr(res, "ErrorCode")) Then
        a = Split(res, "ErrorCode>")(1)
        a = Left(a, Len(a) - 2)
        If (a) Then
            Dim errorPl As String
            Dim errorEN As String
            Dim errorNIP As String
            errorPl = Split(res, "ErrorMessagePl>")(1)
            errorPl = Left(errorPl, Len(errorPl) - 2)
            errorEN = Split(res, "ErrorMessageEn>")(1)
            errorEN = Left(errorEN, Len(errorEN) - 2)
            errorNIP = Split(res, "Nip>")(1)
            errorNIP = Left(errorNIP, Len(errorNIP) - 2)
            
            Range("B" & currentRow).Value = (errorPl + Chr(13) + errorEN + Chr(13) + "NIP: " + errorNIP)
            currentRow = currentRow + 1
            GoTo NextIteration
        End If
    End If
    
    a = Split(res, "DaneSzukajPodmiotyResult>")(1)
    a = Left(a, Len(a) - 2)
    
    Dim sXml As String
    Dim dom As MSXML2.DOMDocument60
    Set dom = New MSXML2.DOMDocument60
    dom.LoadXML a
    Debug.Assert dom.parseError = 0
    Dim xmlSomeCData As MSXML2.IXMLDOMElement
    
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/StatusNip")
    ActiveSheet.Range("B" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Regon")
    ActiveSheet.Range("C" & currentRow).Value = xmlSomeCData.Text
    
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Nazwa")
    Dim nameOut As String
    ''nameOut = Replace(xmlSomeCData.Text, "&", "and")
    nameOut = Replace(xmlSomeCData.Text, "amp;", "")
    'nameOut could be replaced for the cdatatext
    ActiveSheet.Range("D" & currentRow).Value = nameOut
    
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Wojewodztwo")
    ActiveSheet.Range("E" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Powiat")
    ActiveSheet.Range("F" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Gmina")
    ActiveSheet.Range("G" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Miejscowosc")
    ActiveSheet.Range("H" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/KodPocztowy")
    ActiveSheet.Range("I" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/Ulica")
    ActiveSheet.Range("J" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/NrNieruchomosci")
    ActiveSheet.Range("K" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/NrLokalu")
    ActiveSheet.Range("L" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/DataZakonczeniaDzialalnosci")
    ActiveSheet.Range("M" & currentRow).Value = xmlSomeCData.Text
    Set xmlSomeCData = dom.SelectSingleNode("root/dane/MiejscowoscPoczty")
    ActiveSheet.Range("N" & currentRow).Value = xmlSomeCData.Text
    End With
    'Range("C" & currentRow).Value = "some str"
    currentRow = currentRow + 1
    
    ' wyloguj
    With CreateObject("winhttp.winhttprequest.5.1")
            .Open "POST", "https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc", False
            .setRequestHeader "Content-Type", "application/soap+xml;charset=UTF-8;"
            .send "" & _
                    "<soap:Envelope xmlns:soap=""http://www.w3.org/2003/05/soap-envelope"" xmlns:ns=""http://CIS/BIR/PUBL/2014/07"">" & _
                    "<soap:Header xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" & _
                    "<wsa:Action>http://CIS/BIR/PUBL/2014/07/IUslugaBIRzewnPubl/Wyloguj</wsa:Action>" & _
                    "<wsa:To>https://wyszukiwarkaregon.stat.gov.pl/wsBIR/UslugaBIRzewnPubl.svc</wsa:To>" & _
                    "</soap:Header>" & _
                    "<soap:Body>" & _
                    "<ns:Wyloguj>" & _
                    "<ns:pIdentyfikatorSesji>" + id + "</ns:pIdentyfikatorSesji>" & _
                    "</ns:Wyloguj>" & _
                    "</soap:Body>" & _
                    "</soap:Envelope>"
                    
    res = .responseText
    End With
    
NextIteration:
Next
End Sub



Public Function tekstCodePageToCodePage(sStrIn As String, lFromCP As Long, lOutCP As Long) As String

Dim lLenStrOut  As Long
Dim sAscii      As String
Dim lLenAscii   As Long
Dim lCurrentCP   As Long
 
  lCurrentCP = GetACP
 
  If lFromCP = lCurrentCP Then
    sAscii = sStrIn
  Else
    lLenAscii = MultiByteToWideChar(lFromCP, 0&, sStrIn, Len(sStrIn), 0&, 0&)
    sAscii = String$(lLenAscii, vbNullChar)
    lLenAscii = MultiByteToWideChar(lFromCP, 0&, sStrIn, Len(sStrIn), StrPtr(sAscii), lLenAscii)
  End If
 
  If lOutCP = lCurrentCP Then
    tekstCodePageToCodePage = sAscii
  Else
    lLenStrOut = WideCharToMultiByte(lOutCP, 0&, StrPtr(sAscii), Len(sAscii), 0&, 0&, 0&, 0&)
    tekstCodePageToCodePage = String$(lLenStrOut, vbNullChar)
    lLenStrOut = WideCharToMultiByte(lOutCP, 0&, StrPtr(sAscii), Len(sAscii), tekstCodePageToCodePage, lLenStrOut, 0&, 0&)
  End If
 
End Function

