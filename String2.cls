VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mStr As String

Public Function NewInstance(Optional ByVal val As String = "") As String2
    Set NewInstance = New String2
    NewInstance.Value = val
End Function

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
    Value = mStr
End Property

Public Property Let Value(ByVal val As String)
    mStr = val
End Property

Public Property Get Length() As Long
    Length = Len(mStr)
End Property

Public Function CharAt(ByVal position As Long) As String2
    Set CharAt = Me.NewInstance(Mid$(mStr, position + 1, 1))
End Function

Public Function CharCodeAt(ByVal Index As Long) As Long
    If Index < 0 Or Index >= Len(mStr) Then
        CharCodeAt = -1
    Else
        CharCodeAt = AscW(Mid$(mStr, Index + 1, 1))
    End If
End Function

Public Function Concat(ParamArray args() As Variant) As String2
    Dim res As String
    Dim i As Long: For i = LBound(args) To UBound(args)
        res = res & ParseVal(args(i))
    Next
    Set Concat = Me.NewInstance(res)
End Function

Function EndsWith(ByVal searchString As String, Optional ByVal Length As Long = -1) As Boolean
    Dim effectiveLength As Long
    If Length > 0 Then
        effectiveLength = Length
    Else
        effectiveLength = Len(mStr)
    End If

    If Len(searchString) > effectiveLength Then
        EndsWith = False
        Exit Function
    End If

    Dim startPos As Long: startPos = effectiveLength - Len(searchString) + 1
    Dim subString As String: subString = Mid$(mStr, startPos, Len(searchString))

    EndsWith = (subString = searchString)
End Function

Public Function Format(ParamArray args() As Variant) As String2
    ' エスケープ対応
    mStr = VBA.Replace(mStr, "{{", "≪OPEN≫")
    mStr = VBA.Replace(mStr, "}}", "≪CLOSE≫")
    
    ' 正規表現拡張: {index[,alignment][:format]}
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .ignoreCase = False
        .pattern = "\{(\d+)(?:,(-?\d+))?(?:\:([^}]+))?\}"
    End With
    
    Dim matches As Object: Set matches = regex.Execute(mStr)
        
    Dim i As Long: For i = matches.Count - 1 To 0 Step -1
        Dim m As Object: Set m = matches(i)
        Dim idx As Long: idx = CLng(m.SubMatches(0))
        Dim align As Long: align = 0
        If m.SubMatches(1) <> "" Then align = CLng(m.SubMatches(1))
        Dim fmtSpec As String: fmtSpec = m.SubMatches(2)
        
        Dim raw As String
        If idx <= UBound(args) Then
            raw = ParseVal(args(idx))
        Else
            raw = ""
        End If
        
        ' 書式適用
        Dim formatted As String
        If fmtSpec <> "" Then
            On Error Resume Next
            formatted = VBA.Format(raw, fmtSpec)
            If Err.Number <> 0 Then
                Err.Clear
                formatted = raw
            End If
            On Error GoTo 0
        Else
            formatted = raw
        End If
        
        ' パディング（配置）処理
        Dim padWidth As Long: padWidth = Abs(align)
        If padWidth > 0 Then
            If Len(formatted) < padWidth Then
                If align < 0 Then
                    ' 左寄せ
                    formatted = formatted & Space$(padWidth - Len(formatted))
                Else
                    ' 右寄せ
                    formatted = Space$(padWidth - Len(formatted)) & formatted
                End If
            End If
        End If
        
        ' 置換
        mStr = Left$(mStr, m.FirstIndex) & formatted & Mid$(mStr, m.FirstIndex + m.Length + 1)
    Next i
    
    ' エスケープ復元
    mStr = VBA.Replace(mStr, "≪OPEN≫", "{")
    mStr = VBA.Replace(mStr, "≪CLOSE≫", "}")
    
    Set Format = Me.NewInstance(mStr)
End Function

Public Function Includes(ByVal Search As String, Optional ByVal start As Long = 0, Optional ByVal compareMode As VbCompareMethod) As Boolean
    Dim position As Long: position = InStr(start + 1, mStr, ParseVal(Search), compareMode)
    Includes = (position > 0)
End Function

Public Function IndexOf(ByVal Search As String, Optional ByVal start As Long = 0, Optional ByVal compareMode As VbCompareMethod) As Long
    Dim textLength As Long: textLength = Len(mStr)

    If start >= textLength Then
        IndexOf = -1
        Exit Function
    End If

    If start < 0 Then start = 0

    Dim position As Long: position = InStr(start + 1, mStr, ParseVal(Search), compareMode)

    If position > 0 Then
        IndexOf = position
    Else
        IndexOf = -1
    End If
End Function

Public Function Join(ByVal separator As String, ParamArray args()) As String2
    Dim res As String
    Dim i As Long: For i = LBound(args) To UBound(args)
        res = res & ParseVal(args(i))
        If i < UBound(args) Then res = res & separator
    Next
    Set Join = Me.NewInstance(res)
End Function

Public Function LastIndexOf(ByVal Search As Variant, Optional ByVal startPos As Long = -1) As Long
    Search = ParseVal(Search)
    If Len(Search) = 0 Then
        LastIndexOf = -1
        Exit Function
    End If

    Dim maxPos As Long
    If startPos < 0 Or startPos > Len(mStr) - 1 Then
        maxPos = Len(mStr) - Len(Search)
    Else
        maxPos = startPos
    End If

    Dim i As Long: For i = maxPos To 0 Step -1
        If Mid(mStr, i + 1, Len(Search)) = Search Then
            LastIndexOf = i
            Exit Function
        End If
    Next i

    LastIndexOf = -1
End Function

Function Match(ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As String()
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = True
    regex.ignoreCase = ignoreCase

    Dim matches As Object: Set matches = regex.Execute(mStr)
    Dim result() As String
    If matches.Count = 0 Then
        Match = Null
    Else
        ReDim result(matches.Count - 1)
        Dim i As Long: For i = 0 To matches.Count - 1
            result(i) = matches.Item(i).Value
        Next i
        Match = result
    End If
End Function

Public Function PadEnd(ByVal targetLength As Long, Optional ByVal padString As String = " ") As String2
    If targetLength <= Len(mStr) Then
        Set PadEnd = Me.NewInstance(mStr)
        Exit Function
    End If

    If Len(padString) = 0 Then
        padString = " "
    End If

    Dim padLen As Long: padLen = targetLength - Len(mStr)

    Dim repeatedPad As String
    Do While Len(repeatedPad) < padLen
        repeatedPad = repeatedPad & padString
    Loop

    Dim result As String: result = mStr & Left$(repeatedPad, padLen)
    Set PadEnd = Me.NewInstance(result)
End Function

Public Function PadStart(ByVal targetLength As Long, Optional ByVal padString As String = " ") As String2
    If targetLength <= Len(mStr) Then
        PadStart = mStr
        Exit Function
    End If

    If Len(padString) = 0 Then padString = " "

    Dim padLen As Long: padLen = targetLength - Len(mStr)

    Dim repeatedPad As String
    Do While Len(repeatedPad) < padLen
        repeatedPad = repeatedPad & padString
    Loop

    Dim result As String: result = Left$(repeatedPad, padLen) & mStr
    Set PadStart = Me.NewInstance(result)
End Function

Private Function ParseVal(ByVal inputVar As Variant) As String
    On Error GoTo ErrorHandler

    If IsNull(inputVar) Or IsEmpty(inputVar) Then
        ParseVal = ""
        Exit Function
    End If

    If VarType(inputVar) = vbObject Then
        ParseVal = TypeName(inputVar)
        Exit Function
    End If

    ParseVal = CStr(inputVar)
    Exit Function

ErrorHandler:
    ParseVal = "[Error parsing Variant]"
End Function

Public Function Repeat(ByVal Count As Long) As String2
    Dim result As String: result = ""
    
    If Count < 0 Then
        Err.Raise vbObjectError + 513, "repeat", "Repeat count must be non-negative"
    End If
    
    Dim i As Long: For i = 1 To Count
        result = result & mStr
    Next i
    
    Set Repeat = Me.NewInstance(result)
End Function

Public Function Replace(ByVal pattern As String, ByVal replacement As String, Optional ByVal useRegex As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As String2
    If useRegex Then
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = pattern
        regex.Global = False
        regex.ignoreCase = ignoreCase
        Set Replace = Me.NewInstance(regex.Replace(mStr, replacement))
    Else
        Dim pos As Long: pos = InStr(1, mStr, pattern, vbBinaryCompare)
        If pos > 0 Then
            Set Replace = Me.NewInstance(Left$(mStr, pos - 1) & replacement & Mid$(mStr, pos + Len(pattern)))
        Else
            Set Replace = Me.NewInstance(mStr)
        End If
    End If
End Function

Public Function ReplaceAll(ByVal pattern As String, ByVal replacement As String, Optional useRegex As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As String2
    If useRegex Then
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = pattern
        regex.Global = True
        regex.ignoreCase = ignoreCase
        Set ReplaceAll = Me.NewInstance(regex.Replace(mStr, replacement))
    Else
        If ignoreCase Then
            Set ReplaceAll = Me.NewInstance(VBA.Replace(mStr, pattern, replacement, 1, -1, vbBinaryCompare))
        Else
            Set ReplaceAll = Me.NewInstance(VBA.Replace(mStr, pattern, replacement, 1, -1, vbTextCompare))
        End If
    End If
End Function

Public Function Reverse() As String2
    Set Reverse = Me.NewInstance(VBA.StrReverse(mStr))
End Function

Public Function Search(ByVal pattern As String, Optional ignoreCase As Boolean = True) As Long
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = False
    regex.ignoreCase = ignoreCase
    regex.MultiLine = False
    
    If regex.Test(mStr) Then
        Search = regex.Execute(mStr)(0).FirstIndex
    Else
        Search = -1
    End If
End Function

Public Function Slice(Optional ByVal indexStart As Long = 0, Optional ByVal indexEnd As Long = 0) As String2
    Dim startPos As Long
    Dim endPos As Long
    
    If indexStart < 0 Then
        startPos = Me.Length + CLng(indexStart)
    Else
        startPos = CLng(indexStart)
    End If

    If indexEnd <= 0 Then
        endPos = Me.Length + CLng(indexEnd)
    Else
        endPos = CLng(indexEnd)
    End If

    If startPos < 0 Then startPos = 0
    If endPos < 0 Then endPos = 0
    If startPos > Me.Length Then startPos = Me.Length
    If endPos > Me.Length Then endPos = Me.Length

    If endPos <= startPos Then
        Set Slice = Me.NewInstance("")
    Else
        Set Slice = Me.NewInstance(Mid$(mStr, startPos + 1, endPos - startPos))
    End If
End Function

Function Split(Optional ByVal separator As String = "#Empty#", Optional limit As Long = -1) As String()
    Dim result() As String

    Dim useLimit As Boolean: useLimit = False
    Dim maxParts As Long
    If limit >= 0 Then
        If CLng(limit) <= 0 Then
            Split = Array()
            Exit Function
        Else
            maxParts = CLng(limit)
            useLimit = True
        End If
    End If

    If separator = "#Empty#" Then
        ReDim result(0)
        result(0) = mStr
        Split = result
        Exit Function
    End If

    If separator = "" Then
        Dim i As Long
        If useLimit And maxParts < Me.Length Then
            ReDim result(0 To maxParts - 1)
            For i = 1 To maxParts - 1
                result(i - 1) = Mid$(mStr, i, 1)
            Next i
            result(maxParts - 1) = Mid$(mStr, maxParts)
        Else
            ReDim result(0 To Me.Length - 1)
            For i = 1 To Me.Length
                result(i - 1) = Mid$(mStr, i, 1)
            Next i
        End If
        Split = result
        Exit Function
    End If

    Dim sepLength As Long: sepLength = Len(separator)
    Dim startPos As Long: startPos = 1
    Dim Index As Long: Index = 0

    Do
        If useLimit And Index >= maxParts - 1 Then Exit Do
        Dim pos As Long: pos = InStr(startPos, mStr, separator)
        If pos > 0 Then
            ReDim Preserve result(Index)
            result(Index) = Mid$(mStr, startPos, pos - startPos)
            startPos = pos + sepLength
            Index = Index + 1
        Else
            Exit Do
        End If
    Loop

    ReDim Preserve result(Index)
    result(Index) = Mid$(mStr, startPos)

    Split = result
End Function

Public Function StartsWith(ByVal prefix As Variant, Optional position As Long = 0) As Boolean
    prefix = ParseVal(prefix)
    Dim prefixLength As Long: prefixLength = Len(prefix)

    If position < 0 Then
        StartsWith = False
        Exit Function
    End If

    If position >= Me.Length Then
        StartsWith = False
        Exit Function
    End If

    If (Me.Length - position) < prefixLength Then
        StartsWith = False
        Exit Function
    End If

    StartsWith = (Mid$(mStr, position + 1, prefixLength) = prefix)
End Function

Public Function subString(ByVal starts As Long, Optional ByVal ends As Long = -1) As String2
    If ends = -1 Then
        Set subString = Me.NewInstance(Mid$(mStr, starts + 1))
    Else
        Set subString = Me.NewInstance(Mid$(mStr, starts + 1, ends - starts))
    End If
End Function

Public Function ToCharArray() As String()
    Dim arr() As String: ReDim arr(Me.Length - 1)
    Dim i As Long: For i = 0 To Me.Length - 1
        arr(i) = Mid(mStr, i + 1, 1)
    Next
    ToCharArray = arr
End Function

Public Function ToCharCodeArray() As Long()
    Dim arr() As Long: ReDim arr(Me.Length - 1)
    Dim i As Long: For i = 0 To Me.Length - 1
        arr(i) = AscW(Mid$(mStr, i + 1, 1))
    Next
    ToCharCodeArray = arr
End Function

Public Function ToLowerCase() As String2
    Set ToLowerCase = Me.NewInstance(VBA.LCase$(mStr))
End Function

Public Function ToString() As String
    ToString = mStr
End Function

Public Function ToUpperCase() As String2
    Set ToUpperCase = Me.NewInstance(VBA.UCase$(mStr))
End Function

Public Function Trim() As String2
    Set Trim = Me.NewInstance(VBA.Trim(mStr))
End Function

Public Function TrimEnd() As String2
    Set TrimEnd = Me.NewInstance(VBA.RTrim$(mStr))
End Function

Public Function TrimStart() As String2
    Set TrimStart = Me.NewInstance(VBA.LTrim$(mStr))
End Function
