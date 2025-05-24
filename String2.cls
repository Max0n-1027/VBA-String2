VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private mStr As String

Public Property Get length() As Long
    length = Len(mStr)
End Property

Public Function NewInstance(Optional ByVal Val As String = "") As String2
    Set NewInstance = New String2
    NewInstance.Value = Val
End Function

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
    Value = mStr
End Property

Public Property Let Value(ByVal Val As String)
    mStr = Val
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

Public Function Concat(ByVal Val As String) As String2
    Set Concat = Me.NewInstance(mStr & Val)
End Function

Function EndsWith(ByVal searchString As String, Optional ByVal length As Long = -1) As Boolean
    Dim effectiveLength As Long
    If length > 0 Then
        effectiveLength = length
    Else
        effectiveLength = Len(mStr)
    End If

    If Len(searchString) > effectiveLength Then
        EndsWith = False
        Exit Function
    End If

    Dim startPos As Long: startPos = effectiveLength - Len(searchString) + 1
    Dim SubString As String: SubString = Mid$(mStr, startPos, Len(searchString))

    EndsWith = (SubString = searchString)
End Function

Public Function Includes(ByVal searchString As String, Optional ByVal start As Long = 0) As Boolean
    Dim position As Long: position = InStr(start + 1, mStr, searchString, vbBinaryCompare)
    Includes = (position > 0)
End Function

Public Function IndexOf(ByVal searchString As String, Optional ByVal start As Long = 0) As Long
    Dim textLength As Long: textLength = Len(mStr)

    If start >= textLength Then
        IndexOf = -1
        Exit Function
    End If

    If start < 0 Then
        start = 0
    End If

    Dim position As Long: position = InStr(start + 1, mStr, searchString, vbBinaryCompare)

    If position > 0 Then
        IndexOf = position
    Else
        IndexOf = -1
    End If
End Function

Public Function LastIndexOf(Search As String, Optional startPos As Long = -1) As Long
    Dim i As Long
    Dim maxPos As Long

    If Len(Search) = 0 Then
        LastIndexOf = -1
        Exit Function
    End If

    If startPos < 0 Or startPos > Len(mStr) - 1 Then
        maxPos = Len(mStr) - Len(Search)
    Else
        maxPos = startPos
    End If

    For i = maxPos To 0 Step -1
        If Mid(mStr, i + 1, Len(Search)) = Search Then
            LastIndexOf = i
            Exit Function
        End If
    Next i

    LastIndexOf = -1
End Function

Function Match(ByVal pattern As String) As String()
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = True
    regex.ignoreCase = False

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

    ' padString を必要な長さまで繰り返す
    Dim repeatedPad As String
    Do While Len(repeatedPad) < padLen
        repeatedPad = repeatedPad & padString
    Loop

    ' 必要な長さだけ切り取って末尾に追加
    Dim result As String: result = mStr & Left$(repeatedPad, padLen)
    Set PadEnd = Me.NewInstance(result)
End Function

Public Function PadStart(ByVal targetLength As Long, Optional ByVal padString As String = " ") As String2
    If targetLength <= Len(mStr) Then
        PadStart = mStr
        Exit Function
    End If

    If Len(padString) = 0 Then
        padString = " "
    End If

    Dim padLen As Long: padLen = targetLength - Len(mStr)

    ' padString を必要な長さまで繰り返す
    Dim repeatedPad As String
    Do While Len(repeatedPad) < padLen
        repeatedPad = repeatedPad & padString
    Loop

    ' 必要な長さだけ切り取って先頭に追加
    Dim result As String: result = Left$(repeatedPad, padLen) & mStr
    Set PadStart = Me.NewInstance(result)
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

Public Function Replace(ByVal pattern As String, ByVal replacement As String, Optional ByVal useRegex As Boolean = False) As String2
    If useRegex Then
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = pattern
        regex.Global = False  ' 最初の一致のみを置換
        regex.ignoreCase = True
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

Public Function ReplaceAll(ByVal pattern As String, ByVal replacement As String, Optional useRegex As Boolean = False) As String2
    If useRegex Then
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = pattern
        regex.Global = True
        regex.ignoreCase = True
        Set ReplaceAll = Me.NewInstance(regex.Replace(mStr, replacement))
    Else
        Set ReplaceAll = Me.NewInstance(VBA.Replace(mStr, pattern, replacement, 1, -1, vbBinaryCompare))
    End If
End Function

Public Function Reverse() As String2
    Set Reverse = Me.NewInstance(VBA.StrReverse(mStr))
End Function

Public Function Search(ByVal pattern As String) As Long
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.Global = False
    regex.ignoreCase = True
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
        startPos = Me.length + CLng(indexStart)
    Else
        startPos = CLng(indexStart)
    End If

    If indexEnd <= 0 Then
        endPos = Me.length + CLng(indexEnd)
    Else
        endPos = CLng(indexEnd)
    End If

    ' 範囲の調整
    If startPos < 0 Then startPos = 0
    If endPos < 0 Then endPos = 0
    If startPos > Me.length Then startPos = Me.length
    If endPos > Me.length Then endPos = Me.length

    If endPos <= startPos Then
        Set Slice = Me.NewInstance("")
    Else
        Set Slice = Me.NewInstance(Mid$(mStr, startPos + 1, endPos - startPos))
    End If
End Function

Function Split(Optional ByVal separator As String = "#Empty#", Optional limit As Long = -1) As String()
    Dim result() As String
    ' limit の処理
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

    ' separator が省略された場合
    If separator = "#Empty#" Then
        ReDim result(0)
        result(0) = mStr
        Split = result
        Exit Function
    End If

    ' separator が空文字列の場合、1 文字ずつ分割
    If separator = "" Then
        Dim i As Long
        If useLimit And maxParts < Me.length Then
            ReDim result(0 To maxParts - 1)
            For i = 1 To maxParts - 1
                result(i - 1) = Mid$(mStr, i, 1)
            Next i
            result(maxParts - 1) = Mid$(mStr, maxParts)
        Else
            ReDim result(0 To Me.length - 1)
            For i = 1 To Me.length
                result(i - 1) = Mid$(mStr, i, 1)
            Next i
        End If
        Split = result
        Exit Function
    End If

    ' 通常の区切り文字で分割
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

Public Function StartsWith(ByVal prefix As String, Optional position As Long = 0) As Boolean
    Dim prefixLength As Long: prefixLength = Len(prefix)

    ' position が負の場合は false を返す
    If position < 0 Then
        StartsWith = False
        Exit Function
    End If

    ' position が text の長さ以上の場合は false を返す
    If position >= Me.length Then
        StartsWith = False
        Exit Function
    End If

    ' 残りの文字数が prefix より短い場合は false を返す
    If (Me.length - position) < prefixLength Then
        StartsWith = False
        Exit Function
    End If

    ' 指定された位置から prefix の長さ分の文字列を取得して比較
    StartsWith = (Mid$(mStr, position + 1, prefixLength) = prefix)
End Function

Public Function SubString(ByVal starts As Long, Optional ByVal ends As Long = -1) As String2
    If ends = -1 Then
        Set SubString = Me.NewInstance(Mid$(mStr, starts + 1))
    Else
        Set SubString = Me.NewInstance(Mid$(mStr, starts + 1, ends - starts))
    End If
End Function

Public Function ToCharArray() As String()
    Dim arr() As String: ReDim arr(Me.length - 1)
    Dim i As Long: For i = 0 To Me.length - 1
        arr(i) = Mid(mStr, i + 1, 1)
    Next
    ToCharArray = arr
End Function

Public Function ToLowerCase() As String2
    Set ToLowerCase = Me.NewInstance(LCase(mStr))
End Function

Public Function ToString() As String
    ToString = mStr
End Function

Public Function ToUpperCase() As String2
    Set ToUpperCase = Me.NewInstance(UCase(mStr))
End Function

Public Function Trim() As String2
    Set Trim = Me.NewInstance(VBA.Trim(mStr))
End Function

Public Function TrimEnd() As String2
    Dim i As Long
    i = Len(mStr)
    Do While i >= 1 And Mid(mStr, i, 1) = " "
        i = i - 1
    Loop
    mStr = Left(mStr, i)
    Set TrimEnd = Me.NewInstance(mStr)
End Function

Public Function TrimStart() As String2
    Dim i As Long: i = 1
    Do While i <= Len(mStr) And Mid(mStr, i, 1) = " "
        i = i + 1
    Loop
    mStr = Mid(mStr, i)
    Set TrimStart = Me.NewInstance(mStr)
End Function
