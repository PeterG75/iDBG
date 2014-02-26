VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: david zimmer (dzimmer@isightpartners.com dzzie@yahoo.com)
'CopyRight: iSightPartners.com
'License GPL
'Date: Feb 17 2010

Public FunctionName As String
Public retaddr As Long
Public bufsize As Long
Public BufAddr As Long
Public FreedData As String
Public FreedFrom As Long
Public Visible As Boolean
Public li As ListItem
Public NotFound As Boolean


Sub SetLiColor(c As ColorConstants)
    On Error Resume Next
    If Not li Is Nothing Then li.ForeColor = c
End Sub

Function GetReport() As String
    On Error Resume Next
    
    Dim buf() As Byte
    
    Dim ret()
    push ret, "Function: " & FunctionName
    
    If NotFound Then
        push ret, "Matching Alloc Was not found Showing first 50 bytes"
    Else
        push ret, vbTab & "Address: 0x" & Hex(BufAddr)
        push ret, vbTab & "Size: 0x" & Hex(bufsize) & " (" & bufsize & ")"
        push ret, "Alloced @: 0x" & Hex(retaddr)
    End If
    
    push ret, "Free'd @: 0x" & Hex(FreedFrom)
    push ret, vbCrLf
    
    If FreedFrom = 0 Then
        If Len(FreedData) > 0 Then
            If frmLibDebug.dbg.isDebugging Then
                push ret, "Data shown is current memory value. Block has not been freed yet."
            Else
                push ret, "Data was successfully extracted just before termination but was not properly freed"
            End If
        End If
    End If
    
    push ret, "See hex dump pane for true representation of data. Copy and save buttons are pane sensitive"
    push ret, vbCrLf & vbCrLf
    
    push ret, ReadableData
    GetReport = Join(ret, vbCrLf)
    
End Function

Function GetDataBytes() As Byte()
    GetDataBytes = StrConv(FreedData, vbFromUnicode)
End Function

Private Function ReadableData() As String
    
    Dim buf() As Byte
    Dim tmp As String
    Dim x As Byte
    
    If Len(FreedData) = 0 Then
        ReadableData = "[Data not read in yet]"
        Exit Function
    End If
    
    buf = StrConv(FreedData, vbFromUnicode)
    For i = 0 To UBound(buf)
        If i > 200 Then Exit For
        DoEvents
        x = buf(i)
        tmp = tmp & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
    Next
    
    ReadableData = tmp
    
End Function


Private Function HexDump(ary, ByVal offset) As String
    Dim s() As String, chars As String, tmp As String
    
    If offset > 0 And offset Mod 16 <> 0 Then MsgBox "Hexdump isnt being used right! Offset not on boundry"

    'i am lazy and simplicity rules, make sure offset read
    'starts at standard mod 16 boundry or all offsets will
    'be wrong ! it is okay to read a length that ends off
    'boundry though..that was easy to fix...
    
    chars = "   "
    For i = 1 To UBound(ary)
        tmp = tmp & ary(i) & " "
        x = CInt("&h" & ary(i))
        chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            push s, h & "   " & tmp & chars
            offset = offset + 16:  tmp = Empty: chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        h = Hex(offset)
        While Len(h) < 6: h = "0" & h: Wend
        h = h & "   " & tmp
        While Len(h) <= 56: h = h & " ": Wend
        push s, h & chars
    End If
    
    HexDump = Join(s, vbCrLf)
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub
