Option Explicit

Sub test()
    Dim parser As Json2VBA, json As Variant, all_json() As String
    
    Set parser = New Json2VBA
    
    For Each json In json_test_strings()
        Debug.Print CStr(json)
        parser.parse CStr(json)
    Next
End Sub


Function json_test_strings() As String()
    Dim jsonSamples(1 To 12) As String
    ' 1. String inside dictionary
    jsonSamples(1) = "{""key"": ""value""}"
    ' 2. Boolean inside dictionary
    jsonSamples(2) = "{""flag"": true}"
    ' 3. Null inside dictionary
    jsonSamples(3) = "{""nothing"": null}"
    ' 4. Number inside dictionary
    jsonSamples(4) = "{""number"": 123}"
    ' 5. Array inside dictionary
    jsonSamples(5) = "{""items"": [1, 2, 3]}"
    ' 6. Dictionary inside dictionary
    jsonSamples(6) = "{""nested"": {""inner"": ""yes""}}"
    ' 7. String inside array
    jsonSamples(7) = "[""hello""]"
    ' 8. Boolean inside array
    jsonSamples(8) = "[false]"
    ' 9. Null inside array
    jsonSamples(9) = "[null]"
    '10. Number inside array
    jsonSamples(10) = "[42]"
    '11. Array inside array
    jsonSamples(11) = "[[1, 2]]"
    '12. Dictionary inside array
    jsonSamples(12) = "[{""a"": 1}]"
    json_test_strings = jsonSamples
End Function

