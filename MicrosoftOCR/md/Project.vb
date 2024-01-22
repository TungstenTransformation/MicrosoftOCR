'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
'#Language "WWB-COM"
Option Explicit
   'On Menu/Edit/References... Add reference to "Microsoft XML, v 6.0"
   'On Menu/Edit/References... Add reference to "Microsoft VB Regular Expressions 5.5"
   'On Menu/Edit/References... Add reference to "Microsoft Scripting Runtime"

' Project Script

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Kofax Transformation, rename the default page OCR profile to "Microsoft OCR"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="Microsoft OCR" Then MicrosoftOCR_Read(pXDoc)
End Sub

Public Sub MicrosoftOCR_Read(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, JSON As String, P As Long, JS As Dictionary
   While pXDoc.Representations.Count>0
      'If pXDoc.Representations(0).Name="Microsoft OCR" Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   pXDoc.Representations.Create("Microsoft OCR")
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
   For P=0 To pXDoc.CDoc.Pages.Count-1
      'JSON=MicrosoftOCR_REST(pXDoc.CDoc.Pages(P).SourceFileName,EndPoint,Key)
      Open Replace(pXDoc.FileName, ".xdc",".json") For Input As #1
         JSON=Input(LOF(1),1)
      Close 1
      Set JS= JSON_Parse(JSON)
      Exit Sub '  remove XXXXXX! DEBUG
      MicrosoftOCR_AddWords(pXDoc, JS, P, UseMicrosoftTextLines:=True)
   Next
End Sub

Public Function MicrosoftOCR_REST(ImageFileName As String, EndPoint As String, Key As String) As String
   'Call Microsoft Azure Computer Vision OCR API 3.2
   'https://westus.dev.cognitive.microsoft.com/docs/services/computer-vision-v3-2/operations/56f91f2e778daf14a499f20d
   Dim  HTTP As New MSXML2.XMLHTTP60, Image() As Byte
   'supports JPEG, JPG, PNG, TIFF, BMP, 50x50 up to 4200x4200, max 10 megapixel
   Open ImageFileName For Binary Access Read As #1
   ReDim Image (0 To LOF(1)-1)
   Get #1,, Image
   Close #1
   HTTP.Open("POST", EndPoint & "/vision/v3.2/read/analyze",varAsync:=False)
   HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
   HTTP.setRequestHeader("Content-Type", "application/octet-stream")
   HTTP.send(Image)
   Dim getRequestStatus As MatchCollection, RegexAzureStatus As New RegExp, OperationLocation As String, Delay As Long, I As Long, Status As String

   RegexAzureStatus.Pattern = """(?:message|status)"":\s*""(.*?)""" 'Get message or status from JSON via regex
   If HTTP.status<>202 Then
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
      Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & getRequestStatus(0).SubMatches(0))
   End If
   OperationLocation=HTTP.getResponseHeader("Operation-Location") 'Get the URL To retrive the result
   Delay=1 'Wait 1 second for result (Microsoft recommends calling no more frequently than 1 second)
   For I= 1 To 100
      Wait Delay
      Set HTTP = New MSXML2.XMLHTTP60
      HTTP.Open("GET", OperationLocation & "?a=" & CStr(I),varAsync:=False)
      HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
      'HTTP.setRequestHeader("Cache-Control", "no-cache")
      'HTTP.setRequestHeader("Cache-Control", "max-age=0")
      'HTTP.setRequestHeader("cache-control", "private")
      HTTP.send()
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
      Status=getRequestStatus(0).SubMatches(0)
      If HTTP.status<>200 Then Err.Raise (655,,"Microsoft OCR Error: (" & HTTP.status & ") " & Status)
      Select Case Status
      Case "succeeded"
         Exit For
      Case "failed"
            Err.Raise (656,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
      Case "running", "notStarted"
         ' Delay=Delay+1 ' wait 1 second longer next time
      End Select
   Next
   MicrosoftOCR_REST = HTTP.responseText
End Function

Public Sub MicrosoftOCR_AddWords(pXDoc As CscXDocument, JS As Object, PageOffset As Long, Optional UseMicrosoftTextLines As Boolean)
   Dim P As Long, W As Long, Key As String, Confidences As String, Word As CscXDocWord, Units As String, XRes As Double, YRes As Double, L As Long
   For P=0 To JS("js.analyzeResult.readResults._count")-1
      Units=JS("js.analyzeResult.readResults(" & CStr(P) & ").unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(P).XRes
         YRes=pXDoc.CDoc.Pages(P).XRes
      End If
      For L=0 To JS("js.analyzeResult.readResults(" & P & ").lines._count")-1   'format
          For W=0 To JS("js.analyzeResult.readResults(" & P & ").lines(" & L & ").words._count")-1   'format
             Key="js.analyzeResult.readResults(" & P & ").lines(" & L & ").words(" & W & ")"
             Set Word = New CscXDocWord
             Word.Text=JSON_Unescape(JS(Key & ".text"))
             Word.PageIndex=P
             If UseMicrosoftTextLines Then 'Give all the words FAKE coordinates so that KT sees Microsoft's Textlines
               Word.Left=W*10
               Word.Width=5
               Word.Top=L*10
               Word.Height=5
             Else
                JSON_Polygon2Rectangle(JS,Key,Word,Units,XRes,YRes) 'Give the words the correct coordinates
             End If
             Confidences = Confidences & JS(Key & ".confidence") & ","
             pXDoc.Pages(P+PageOffset).AddWord(Word)
          Next
      Next
   Next
   Confidences = Left(Confidences,Len(Confidences)-1) ' trim trailing ,
   'Store all confidences for later use in AZL
   If pXDoc.XValues.ItemExists("MicrosoftOCR_WordConfidences") Then pXDoc.XValues.Delete("MicrosoftOCR_WordConfidences")
   pXDoc.XValues.Add("MicrosoftOCR_WordConfidences",Confidences,True)
   pXDoc.Representations(0).AnalyzeLines 'Redo Text Line Analysis in Kofax Transformation
   If Not UseMicrosoftTextLines Then Exit Sub
   'restore word coordinates after textlines created
   For W=0 To pXDoc.Words.Count-1
      Set Word=pXDoc.Words(W)
      Key="js.analyzeResult.readResults(" & Word.PageIndex & ").lines(" & Word.LineIndex & ").words(" & Word.IndexInTextLine & ")"
      Units=JS("js.analyzeResult.readResults(" & CStr(Word.PageIndex) & ").unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(P).XRes
         YRes=pXDoc.CDoc.Pages(P).XRes
      End If
      JSON_Polygon2Rectangle(JS,Key,Word,Units,XRes,YRes)
   Next
End Sub

Public Function min(A,b)
   If A<b Then min=A Else min=b
End Function
Public Function max(A,b)
   If A>b Then max=A Else max=b
End Function

Public Sub JSON_Polygon2Rectangle(JS As Object, Key As String, Rectangle As Object, Units As String, XRes As Long, YRes As Long)
   'Microsoft returns the coordinates of a region as JSON ->   "polygon": [1848,492,1896,494,1897,535,1849,535]
   'We need to convert this to  .left, .width, .top and .height
   Rectangle.Left=  min(JS(Key & ".boundingBox(0)"),JS(Key & ".boundingBox(6)"))
   Rectangle.Width= max(JS(Key & ".boundingBox(2)"),JS(Key & ".boundingBox(4)"))-Rectangle.Left
   Rectangle.Top =  min(JS(Key & ".boundingBox(1)"),JS(Key & ".boundingBox(3)"))
   Rectangle.Height=max(JS(Key & ".boundingBox(5)"),JS(Key & ".boundingBox(7)"))-Rectangle.Top
   If Units="inch" Then
      Rectangle.Left=Rectangle.Left*XRes
      Rectangle.Width=Rectangle.Width*XRes
      Rectangle.Top=Rectangle.Top*YRes
      Rectangle.Height=Rectangle.Height*YRes
   End If
End Sub


'-------------------------------------------------------------------
' VBA JSON Parser https://github.com/KofaxTransformation/KTScripts/blob/master/JSON%20parser%20in%20vb.md
'-------------------------------------------------------------------
Private T As Long, Tokens As Object, Keys As Object
Function JSON_Parse(JSON As String, Optional Key As String = "$") As Object
   'This is 100% compliant with ECMA-404 JSON Data Interchange Standard at https://www.json.org/json-en.html
   'the regex pattern finds strings including characters escaped with \ OR numbers OR true/false/null OR \\{}:,[]
   'tested at https://regex101.com/r/YkiVdc/1
   'This script will crash on invalid JSON
   With CreateObject("vbscript.regexp")
      .Global=True
      .Pattern = """(?:[^""\\]|\\.)*""|-?(?:\d+)(?:\.\d*)?(?:[eE][+\-]?\d+)?|(?:true|false|null)|[\[\]{}:,]"
      Set tokens=.Execute(JSON)
   End With
   T = 0
   Set Keys = CreateObject("Scripting.Dictionary")
   If Tokens(T) = "{" Then JSON_ParseObject(Key) Else JSON_ParseArray(Key)
   Return Keys
End Function

Sub JSON_ParseObject(Key As String)
    Do
      T = t + 1
     Select Case tokens(t).Value
         Case "]"
         Case "[":  JSON_ParseArray(Key)
         Case "{"
                    If tokens(t + 1).Value = "}" Then
                        t = t + 1
                        Keys.Add(Key, Nothing) 'empty object
                    Else
                        JSON_ParseObject(Key)
                    End If
         Case "}":
            If tokens(t - 1).Value = "{" Then Keys.Add(Key, Nothing) ' this was an empty object
            Key = JSONPath_Parent(Key)
            Exit Do
         Case ":":  Key = Key & "." & JSON_Value(tokens(t - 1).Value) 'previous token was a name - remember it
         Case ",":  Key = JSONPath_Parent(Key)
         Case Else 'we are in a string. if next is not ":" then we are value
            If tokens(t + 1).Value <> ":" Then Keys.Add(Key, JSON_Value(tokens(t).Value))
     End Select
     Keys_ToText
    Loop
End Sub

Sub JSON_ParseArray(Key As String)
   Dim A As Long ' Array index
   Do
      T = t + 1
      Select Case tokens(t).Value
         Case "[":
            If tokens(t + 1).Value = "]" Then 'empty array
               Key=Key & JSON_ArrayID(A)
               Keys.Add(Key, Nothing)
               A=-1
               Exit Do
            End If
            JSON_ParseArray(Key & JSON_ArrayID(A)) 'start of an array inside an array
         Case "{":
            JSON_ParseObject(Key & JSON_ArrayID(A)) 'it's an object so recurse
         Case "}":
            If tokens(t - 1).Value = "}" Then Keys.Add(Key & JSON_ArrayID(A), Nothing) 'empty object
         Case ":":  Key = Key & JSON_ArrayID(A)
         Case ",":  A = A + 1
         Case "]":
            Exit Do
         Case Else: Keys.Add(Key & JSON_ArrayID(A), JSON_Value(tokens(t).Value))
      End Select
      Keys_ToText
   Loop
   Keys.Add(Key & ".length",A + 1) 'store array length in dictionary
   Keys_ToText
End Sub

Function JSON_ArrayID(e As Long) As String
    Return "(" & CStr(e) & ")"
End Function

Function JSONPath_Parent(Key As String) As String 'go to the parent key
    If InStr(Key, ".") Then Return Left(Key, InStrRev(Key, ".") - 1)
End Function

Function JSON_Value(Value As String) 'JSON values can be string, number, true, false or null
   'Strings start with a " in JSON - everything else is true,false, null or a number
   Dim Locale As Long, Number As Double
   If Left (Value,1)="""" Then Return JSON_Unescape(Mid(Value,2,Len(Value)-2)) 'strip " from begin and end of string
   Select Case Value
   Case "true"
      Return True
   Case "false"
      Return False
   Case "null"
      Return Nothing
   Case Else 'it has to be a number
      Locale=GetLocale() 'preserve locale
      SetLocale(1033) 'en_us
      'these are valid JSON numbers: 1 -1 0 -0.1 1111111111 0.1 1.0000 1.0e5 -1e-5 1E5 0e3 0e-3
      'these are invalid JSON numbers, but CDbl converts them correctly: +1 .6 1.e5 -.5 e6
      Number=CDbl(Value) 'CDbl() function luckily converts all allowed JSON number formats
      SetLocale(Locale)
      Return Number
   End Select
End Function

Public Function JSON_Unescape(A As String) As String
   'https://www.json.org/json-en.html
   Dim Hex As String
   A=Replace(A,"\""","""") 'double quote
   A=Replace(A,"\/","/") 'forward slash
   A=Replace(A,"\b",vbBack) 'backspace
   A=Replace(A,"\f",vbLf) 'form feed
   A=Replace(A,"\n",vbCrLf) 'new line
   A=Replace(A,"\r",vbCr) 'carraige return
   A=Replace(A,"\t",vbTab) 'tab
   A=Replace(A,"\\","\") 'backslash
   While InStr(A,"\u")  'hex encoded Unicode characters
      Hex=Mid(A,InStr(A,"\u")+2,4)
      A=Replace(A,"\u" & Hex, ChrW(Val("&H" & Hex)))
   Wend
   'This is not handling \u with 4 hex digits
   Return A
End Function

Sub Keys_ToText()  'for Debugging
   Dim Key As String, Value As Variant
   Open "c:\temp\keys.txt" For Output As #1
   Print #1, vbUTF8BOM;
   For Each Key In Keys.Keys
      Value=Keys(Key)
      If TypeName(Value)="String" Then Value = """" & Value & """"
      If TypeName(Value)="Nothing" Then Value = "null"
      Print #1, Key & " : " & Value
   Next
   Close 1
End Sub
