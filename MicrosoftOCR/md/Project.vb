'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
'#Reference {BEE4BFEC-6683-3E67-9167-3C0CBC68F40A}#2.4#0#C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.tlb#System#System
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
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
   Dim EndPoint As String, Key As String, JSON As String, P As Long, JS As Object, JSONFile As String
   While pXDoc.Representations.Count>0
      'If pXDoc.Representations(0).Name="Microsoft OCR" Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
   For P=0 To pXDoc.CDoc.Pages.Count-1
      JSONFile = Replace(pXDoc.FileName, ".xdc", "_" & Format("00",P) & ".json")
      If Dir(JSONFile)="" Then
         JSON=MicrosoftOCR_REST(pXDoc.CDoc.Pages(P).SourceFileName,EndPoint,Key)
         Open JSONFile For Output As #1
            Print #1, vbUTF8BOM & JSON
         Close 1
      Else
         Open JSONFile For Input As #1
            JSON=Input(LOF(1),1)
         Close 1
      End If
      Set JS= JSON_Parse(JSON)
      If pXDoc.Representations.Count=0 Then pXDoc.Representations.Create("Microsoft OCR")
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
   Dim P As Long, W As Long, Confidences As String, Word As CscXDocWord, Units As String, XRes As Double, YRes As Double, L As Long
   Dim readResults As Object, ocrWord As Object
   Set readResults=JS("analyzeResult")("readResults")
   For P=0 To readResults.Count-1
      Units=readResults(P)("unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(P).XRes
         YRes=pXDoc.CDoc.Pages(P).XRes
      End If
      For L=0 To readResults(P)("lines").Count-1   'format
         For W=0 To readResults(P)("lines")(L)("words").Count-1   'format
            Set ocrWord = readResults(P)("lines")(L)("words")(W)
             Set Word = New CscXDocWord
             Word.Text=ocrWord("text")
             Word.PageIndex=P
             If UseMicrosoftTextLines Then 'Give all the words FAKE coordinates so that KT sees Microsoft's Textlines
               Word.Left=W*10
               Word.Width=5
               Word.Top=L*10
               Word.Height=5
             Else
                BoundingBox2Rectangle(ocrWord("boundingBox"),Word,Units,XRes,YRes) 'Give the words the correct coordinates
             End If
             Confidences = Confidences & Format("0.000", ocrWord("confidence")) & ","
             pXDoc.Pages(P+PageOffset).AddWord(Word)
         Next 'Word
      Next 'Line
   Next 'Page
   Confidences = Left(Confidences,Len(Confidences)-1) ' trim trailing ,
   'Store all confidences for later use in AZL
   If pXDoc.XValues.ItemExists("MicrosoftOCR_WordConfidences") Then pXDoc.XValues.Delete("MicrosoftOCR_WordConfidences")
   pXDoc.XValues.Add("MicrosoftOCR_WordConfidences",Confidences,True)
   pXDoc.Representations(0).AnalyzeLines 'Redo Text Line Analysis in Kofax Transformation
   If Not UseMicrosoftTextLines Then Exit Sub
   'restore word coordinates after textlines created
   For W=0 To pXDoc.Words.Count-1
      Set Word=pXDoc.Words(W)
      Set ocrWord = readResults(Word.PageIndex)("lines")(Word.LineIndex)("words")(Word.IndexInTextLine)
      Units=readResults(Word.PageIndex)("unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(P).XRes
         YRes=pXDoc.CDoc.Pages(P).XRes
      End If
      BoundingBox2Rectangle(ocrWord("boundingBox"),Word,Units,XRes,YRes)
   Next
End Sub

Public Sub BoundingBox2Rectangle(bb As Object, Rectangle As Object, Units As String, XRes As Long, YRes As Long)
   'Microsoft returns the coordinates of a region as JSON ->   "polygon": [1848,492,1896,494,1897,535,1849,535]
   'We need to convert this to  .left, .width, .top and .height
   Rectangle.Left=  min(bb(0),bb(6))
   Rectangle.Width= max(bb(2),bb(4))-Rectangle.Left
   Rectangle.Top =  min(bb(1),bb(3))
   Rectangle.Height=max(bb(5),bb(7))-Rectangle.Top
   If Units="inch" Then
      Rectangle.Left=Rectangle.Left*XRes
      Rectangle.Width=Rectangle.Width*XRes
      Rectangle.Top=Rectangle.Top*YRes
      Rectangle.Height=Rectangle.Height*YRes
   End If
End Sub

Public Function min(A,b)
   If A<b Then min=A Else min=b
End Function
Public Function max(A,b)
   If A>b Then max=A Else max=b
End Function




'-------------------------------------------------------------------
' VBA JSON Parser https://github.com/KofaxTransformation/KTScripts/blob/master/JSON%20parser%20in%20vb.md
'-------------------------------------------------------------------
Private T As Long, Tokens As Object
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
   T=0
   Select Case Tokens(0)
      Case "{"  : Return JSON_ParseObject()
      Case "["  : Return JSON_ParseArray()
      Case Else : Return JSON_Value(tokens(0))  'Yes a JSON may contain just 1 value
   End Select
End Function

Function JSON_ParseObject() As Object
   Dim Obj As Object, n As String 'Objects contained named objects, arrays or values
   Set Obj = CreateObject("Scripting.Dictionary")
   If tokens(t+1)="}" Then  T=T+2 : Return Obj ' empty object
   Do
      T = t + 1
      Select Case tokens(t).Value
         Case "{"  :  Obj.Add(n,JSON_ParseObject())
         Case "["  :  Obj.Add(n,JSON_ParseArray())
         Case ":"  :  n = JSON_Value(tokens(t-1))
         Case ","
         Case "}"  :  Return Obj
         Case Else : If tokens(t - 1) = ":" Then Obj.Add(n, JSON_Value(tokens(t)))
      End Select
   Loop
End Function

Function JSON_ParseArray()
   Dim A As Object 'Declare A as an array of anything - it may contain strings, booleans, numbers, objects and arrays
   Set A=CreateObject("System.Collections.Sortedlist")
   If tokens(t+1)="]" Then : T=T+2 : Return A ' empty array
   Do
      T = t + 1
      Select Case tokens(t)
         Case "{"  : A.Add(A.Count,JSON_ParseObject()) 'it's an object so recurse
         Case "["  : A.Add(A.Count,JSON_ParseArray()) 'start of an array inside an array
         Case ","  :
         Case "]"  : Return A
         Case Else : A.Add(A.Count,JSON_Value(tokens(t)))
      End Select
   Loop
End Function

Function JSON_Value(Value As String) 'JSON values can be string, number, true, false or null
   'Strings start with a " in JSON - everything else is true,false, null or a number
   If Left (Value,1)="""" Then Return JSON_Unescape(Mid(Value,2,Len(Value)-2)) 'strip " from begin and end of string
   Select Case Value
      Case "true"  : Return True
      Case "false" : Return False
      Case "null"  : Return Nothing
      Case Else 'it has to be a number
         Dim Locale As Long, Number As Double
         Locale=GetLocale() 'preserve locale
         SetLocale(1033) 'en_us
         'these are valid JSON numbers: 1 -1 0 -0.1 1111111111 0.1 1.0000 1.0e5 -1e-5 1E5 0e3 0e-3
         'these are invalid JSON numbers, but CDbl converts them correctly: +1 .6 1.e5 -.5 e6
         Number=CDbl(Value) 'CDbl() function luckily correctly converts all allowed JSON number formats
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
   Return A
End Function
