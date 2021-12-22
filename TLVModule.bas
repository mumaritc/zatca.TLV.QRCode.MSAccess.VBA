Attribute VB_Name = "TLVModule"
Option Explicit
Option Compare Database

Function Tlv(sellersName As String, VatNumber As String, TimeStamp As String, InvoiceTotal As String, VATTotal As String) As String
    
    Dim tags(4), values(4) As String
        tags(0) = 1
        tags(1) = 2
        tags(2) = 3
        tags(3) = 4
        tags(4) = 5
        values(0) = sellersName
        values(1) = VatNumber
        values(2) = TimeStamp
        values(3) = InvoiceTotal
        values(4) = VATTotal
    
    Dim bytes() As Byte
    
    ReDim bytes(0)
    
    Dim v As Integer
    
    For v = LBound(values) To UBound(values)
        
        Dim thisB() As Byte
        thisB() = Utf8BytesFromString(values(v))
        
        Dim leng As Integer
        leng = UBound(thisB) - LBound(thisB) + 1
        
            If UBound(bytes) = 0 Then
                bytes(0) = CByte(tags(v))
            Else
                AppendByte bytes, CByte(tags(v))  'Tag
            End If
        AppendByte bytes, CByte(leng) 'Length
        AppendBytes bytes, thisB        'Value
    Next
    
    Dim base64 As String
    
    Dim z As Integer, strHexOutput As String
    strHexOutput = ""
    
    For z = 0 To UBound(bytes)
    '    strHexOutput = strHexOutput & Hex(bytes(z)) '& vbNewLine
        strHexOutput = strHexOutput & Replace(Space(2 - Len(Hex(bytes(z)))), " ", "0") & Hex(bytes(z)) '& vbNewLine
    Next
    
    Form_frmTLV.HexCode.Value = strHexOutput 'Output HEXCODE of string to TLV form
    strHexOutput = HexToB64(strHexOutput)
    Form_frmTLV.QRCode.Value = strHexOutput 'Output Base64 of HEXCODE to TLV form
    Tlv = strHexOutput
    'Debug.Print strHexOutput
    
End Function

Private Function AppendByte(ByRef arr() As Byte, ByRef b As Byte)
    Dim bytes(0) As Byte
    bytes(0) = b
    AppendBytes arr, bytes
End Function

Private Function AppendBytes(ByRef arr() As Byte, ByRef bytesToAppend() As Byte)
    Dim orgSize, appendSize, i As Integer
    orgSize = UBound(arr) - LBound(arr) + 1
    appendSize = UBound(bytesToAppend) - LBound(bytesToAppend) + 1
    ReDim Preserve arr(orgSize + appendSize - 1)
    For i = 0 To appendSize - 1
        arr(orgSize + i) = bytesToAppend(i)
    Next
End Function

Function HexToB64(ByVal strContent)
  Dim i, c, strReturned, b64map, b64pad, intLen
  b64map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  b64pad = "="
  intLen = Len(strContent)
  For i = 0 To intLen - 3 Step 3
      c = CLng("&h" & Mid(strContent, i + 1, 3))
      strReturned = strReturned & Mid(b64map, Int(c / 64 + 1), 1) & Mid(b64map, (c And 63) + 1, 1)
  Next
 
  If i + 1 = intLen Then
     c = CLng("&h" & Mid(strContent, i + 1, 1))
     strReturned = strReturned & Mid(b64map, c * 4 + 1, 1)
  ElseIf i + 2 = intLen Then
     c = CLng("&h" & Mid(strContent, i + 1, 2))
     strReturned = strReturned & Mid(b64map, Int(c / 4) + 1, 1) & Mid(b64map, (c And 3) * 16 + 1, 1)
  End If
 
  While (Len(strReturned) And 3) > 0
      strReturned = strReturned & b64pad
  Wend
  HexToB64 = strReturned
End Function

