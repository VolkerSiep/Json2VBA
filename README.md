# Json2VBA
A fast and ruthless json parser for VBA

## What is this, another json parser?
Despite what everybody says, I had a hard time finding a json parser in VBA for Excel that suits my needs. Sure, there is the famous VBA-JSON by Tim Hall (https://github.com/VBA-tools/VBA-JSON), but though being complete and a good library, it turned out to cause a noticable performance impact for my applications. Then there is VBA-JSON-parser by omegastripes (https://github.com/omegastripes/VBA-JSON-parser) with a lot of nice additional functionality, but while I'm really a fan of copyleft software, I'm also needing this for development within the industry I work, so at least in this case, I cannot use components licensed as GPLv3.

While looking around, I also found this blog by Daniel Ferry (https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a), and honestly, I got the approach of using regular expressions from here, and consequently the idea to first tokenize the entire json file. But I wished to preserve the data types.

**So what did I do?** I hereby implemented a regular expression based json parser that assumes valid json and only does a minimum of syntax interpretation to extract the data at hand. This means for instance: The parser will parse just fine the following

    {"key1": "value", "key2": [1.2, 3.14, 5.3e-3], "key3": [true, false, null]}

But the above can be modified with the same result to for instance

    {"key1" $ "value" Ø "key2" @ [Ops1.2y, 3.14=Pi, 5.3e-3], "key3": [true really!, false maybe, nullification]}

This is by design because I decide not to use any time on validation, as I trust my json-rpc server to get the format right. Needless to say, I refer to the NO-WARRANTY clausel in the used MIT license terms. 

**For parsing 12000 json strings in average as large as in the `test2`, Json2VBA uses 0.5 seconds, while for instance VBA-JSON required 7 seconds.**

**In conclusion: If you need something slow that also validates the incoming json, this library is not for you.**

## Interface and usage

A minimal test looks as follows:

    Sub test()
        Dim parser As Json2VBA
        Set parser = New Json2VBA ' object can/should be reused to parse many json strings
        Debug.Print parser.parse("3.14159")  ' I know, this is impressive, right!? :->
    End Sub

I guess in most applications, the incoming data might be a dictionary on top level, for instance when parsing json-rpc responses, such as

    {"result": [1855085.7631189728, 0.07119041355365928, 149.29809737863215, 10.628361951949035, 2.4176595145330218e-05, 0.06824596008789391], "id": "0", "jsonrpc": "2.0"}
    
In this case, the returned data type is an object, and thanks to the design of VBA, one must use the `Set` keyword:

    Sub test2()
        ...
        Dim json As String
        json = "{""result"": [1855085.7631189728, 0.07119041355365928, 149.29809737863215, " & _
                             "10.628361951949035, 2.4176595145330218e-05, 0.06824596008789391], " & _
               """id"": ""0"", ""jsonrpc"": ""2.0""}"
        Set result = parser.parse(json)
        Debug.Print result("result")(5) ' -> 2.4176595145330218e-05
    End Sub
    
Normally, you will know what will come back, but if not, there is a dedicated method to tell whether the json string yields an object:

    Sub test3()
        ...
        Debug.Print parser.is_object("{}") ' -> True
        Debug.Print parser.is_object("[]") ' -> False
        Debug.Print parser.is_object("""Just a string""") ' -> False
        Debug.Print parser.is_object(3.14159) ' -> False
    End Sub

The returned data structure is naturally nested - according to the json string. The mapping to VBA data types is as follows:

| Json datatype | VBA datatype         |
|---------------|----------------------|
| string        | String               |
| any number    | Double               |
| boolean       | Boolean              |
| null          | Null                 |
| array         | Array                |
| mapping       | Scripting.Dictionary |

Dictionary keys must be strings, BTW.

## Installation

You only really need the `Json2VBA.cls` file and import it into your Excel project, either a macro-enabled workbook or an Add-In file. That file contains the MIT license terms and a link to the hosting github repository. Please leave it there! The `LICENSE.TXT` file from the repo doesn't need to follow the code. Then, from the "Tools" menue within VBA editor, select "References" and enable the following entries:

| Reference name                             | Why do I need that?                                                    |
|--------------------------------------------|------------------------------------------------------------------------|
| Microsoft VBScript Regular Expressions 5.5 | Used by Json2VBA to efficiently convert the input to a list of tokens  |
| Microsoft Scripting Runtime                | To be able to create dictionary structures (`Scripting.Dictionary`)    |

## Support and questions

As you might have noticed, this project is not (yet) suffering from an overwhelming user base. If you have a question, suggestion or otherwise any comment, please just use the issue tracker to create a new issue.
I will probably read and react to it soon after.

## Developments and Contributions

I like this to be small and consise, and most of all: performant. And as it is, it should cover the intended needs.
Contributions are welcome, don't get me wrong. Just note that I would not like to compromise performance, but for instance rather add another independent method to validate json.
Json encoding is also missing and could be added. I personally do not need a generic json encoder, but use specific code for my data types at hand (again for performance).
