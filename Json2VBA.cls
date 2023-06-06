VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Json2VBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Json Parser
' ===========
' Github: ....
'
' Short documentation
' ===================
' This class is developed for fast but lightweight json parsing in VBA, as existing code is either slow, incomplete, has a large footprint,
' or their license model prevent use in certain (industrial) settings.
' As this code is optimised for speed, it is designed to parse valid json fast, while it's not checking the validity of the format beyond what's required
' to parse the string. For instance, the valid json string
'
' {"key1": "value", "key2": [1.2, 3.14, 5.3e-3], "key3": [true, false, null]}
'
' can be modified with the same result to
'
' {"key1" $ "value" Ø "key2" @ [Ops1.2y, 3.14=Pi, 5.3e-3], "key3": [true really!, false maybe, nullification]}
'
' If you require a slow but validating JSON parser, you came to the wrong code.
'
'
' This class module is licenced under the MIT License
' ===================================================
'
' Copyright 2023 Volker Siepmann
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”),
' to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
' OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit

Private regex_token As RegExp
Private matches As Object
Private current_token As Integer
Private content As Variant
Private dec_sep As String

' ***  Public interface

Public Function parse(ByRef json As String) As Variant
    ' Parse a json string. The mapping of types is as follows:
    '   string     -> string
    '   any number -> double
    '   true|false -> boolean
    '   null       -> Null
    '   array      -> VBA array
    '   dictionary -> Scripting.Dictionary
    '
    ' Note:
    '   The return type might be an object or not.
    '   To allow the client code to use the correct assigment syntax,
    '   call the "is_object" method in this class with the same json string.
    Dim token As String
    tokenize_json json
    token = fetch_token
    ' need the following branching because of VBA's stupid object assignments
    If Left(token, 1) = "{" Then
        Set parse = parse_dict
    Else
        parse = parse_next(token)
    End If
End Function

Public Function is_object(ByRef json As String) As Boolean
    ' returns whether the parsed json string will be an object or not.
    ' this can be used to either SET the result or use normal assignment.
    is_object = (Left(json, 1) = "{")
End Function

' *** End of public interface

Private Sub class_initialize()
    Const RE_STRING As String = """(\\""|[^""])*"""
    Const RE_NUMBER As String = "[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?"
    Const RE_OTHER As String = "(true|false|null)"
    Const RE_TOKEN As String = "(" & RE_STRING & "|" & RE_NUMBER & "|" & RE_OTHER & "|[\[\]{}])"

    Set regex_token = New RegExp
    regex_token.Pattern = RE_TOKEN
    regex_token.Global = True
    dec_sep = Mid(Format(1000, "#,##0.00"), 6, 1) ' this works while Application.DecimalSeparator doesn't
End Sub

Private Function parse_next(token As String) As Variant
    Select Case Left(token, 1)
    Case "{"
        Set parse_next = parse_dict
    Case "["
        parse_next = parse_array
    Case """"
        parse_next = parse_string(token)
    Case "t"
        parse_next = True
    Case "f"
        parse_next = False
    Case "n"
        parse_next = Null
    Case Else ' must be a number
        parse_next = cdbl2(token)
    End Select
End Function

Private Function cdbl2(value As Variant) As Variant
    If dec_sep = "," Then
        value = Replace(value, ".", ",")
    End If
    cdbl2 = CDbl(value)
End Function

Private Function parse_string(token As String) As String
    parse_string = Replace(Replace(Mid(token, 2, Len(token) - 2), "\n", vbCrLf), "\t", vbTab)
End Function


Private Function parse_dict() As Scripting.Dictionary
    Dim token As String
    Dim result As Scripting.Dictionary
    
    Set result = New Scripting.Dictionary
    While True
        token = fetch_token
        If token = "}" Then
            Set parse_dict = result
            Exit Function
        End If
        result.Add parse_string(token), parse_next(fetch_token)
    Wend
End Function

Private Function parse_array() As Variant
    Dim num As Long
    Dim token As String
    Dim result() As Variant ' is collection faster?
    num = 0
    While True
        token = fetch_token
        If token = "]" Then
            If num > 0 Then
                ReDim Preserve result(1 To num)
            End If
            parse_array = result
            Exit Function
        End If
        num = num + 1
        If (Not Not result) = 0 Then ' test whether result is not yet dimensioned
            ReDim result(1 To 32)
        ElseIf num > UBound(result) Then
            ReDim Preserve result(1 To 2 * (num - 1))
        End If
        result(num) = parse_next(token)
    Wend
End Function

Private Sub tokenize_json(ByRef json As String)
    Set matches = regex_token.Execute(json)
    current_token = 0
End Sub

Private Function fetch_token() As String
    fetch_token = matches(current_token).value
    current_token = current_token + 1
End Function
