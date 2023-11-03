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
      If pXDoc.Representations(0).Name="Microsoft OCR" Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   pXDoc.Representations.Create("Microsoft OCR")
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
   For P=0 To pXDoc.CDoc.Pages.Count-1
      JSON=MicrosoftOCR_REST(pXDoc.CDoc.Pages(P).SourceFileName,EndPoint,Key)
      Set JS= JSON_Parse(JSON)
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

'-------------------------------------------------------------------
' VBA JSON Parser https://github.com/KofaxTransformation/KTScripts/blob/master/JSON%20parser%20in%20vb.md
'-------------------------------------------------------------------
Private t As Long, tokens() As String, dic As Object
Function JSON_Parse(JSON$, Optional Key$ = "js") As Object
    t = 1
    tokens = JSON_Tokenize(JSON)
    Set dic = CreateObject("Scripting.Dictionary")
    If tokens(t) = "{" Then JSON_ParseObj(Key) Else JSON_ParseArr(Key)
    Return dic
End Function
Function JSON_ParseObj(Key$)
    Do
      t = t + 1
     Select Case tokens(t)
         Case "]"
         Case "[":  JSON_ParseArr(Key)
         Case "{"
                    If tokens(t + 1) = "}" Then
                        T = T + 1
                        dic.Add(Key, "null")
                    Else
                        JSON_ParseObj(Key)
                    End If

         Case "}":  Key = JSON_ParentPath(Key): Exit Do
         Case ":":  Key = Key & "." & tokens(T - 1) 'previous token was a key - remember it
         Case ",":  Key = JSON_ParentPath(Key)
         Case Else 'we are in a string. if next is not ":" then we are value - so add to dict!
            If tokens(T + 1) <> ":" Then dic.Add(Key, tokens(T))
     End Select
    Loop
End Function
Function JSON_ParseArr(Key$)
   Dim A As Long
   Do
      T = T + 1
      Select Case tokens(T)
         Case "}"
         Case "{":  JSON_ParseObj(Key & JSON_ArrayID(A))
         Case "[":  JSON_ParseArr(Key)
         Case "]":  Exit Do
         Case ":":  Key = Key & JSON_ArrayID(A)
         Case ",":  A = A + 1
         Case Else: dic.Add(Key & JSON_ArrayID(A), tokens(T))
      End Select
   Loop
   dic.Add(Key & "._count",A+1) 'store array length in dictionary
End Function

Function JSON_Tokenize(S As String) 'completely split the JSON string fast into an array of tokens for the parsers
   Dim C As Long, m As Object, n As Object, tokens() As String
   Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
   With CreateObject("vbscript.regexp")
      .Global = True
      .Multiline = False
      .IgnoreCase = True
      .Pattern = Pattern
      Set m = .Execute(S)
      ReDim tokens(1 To m.Count)
      For Each n In m
        C = C + 1
        tokens(C) = n.Value
        If True Then ' bGroup1Bias=?? when is this needed
           If Len(n.SubMatches(0)) Or n.Value = """""" Then
              tokens(C) = n.SubMatches(0)
           End If
        End If
      Next
  End With
  Return tokens
End Function

Function JSON_ArrayID(e) As String
    Return "(" & e & ")"
End Function

Function JSON_ParentPath(Key As String) As String 'go to the parent key
    If InStr(Key, ".") Then Return Left(Key, InStrRev(Key, ".") - 1)
    'else?
End Function

Public Function JSON_Unescape(A As String) As String
   'https://www.json.org/json-en.html
   A=Replace(A,"\""","""") 'double quote
   A=Replace(A,"\\","\") 'backslash
   A=Replace(A,"\/","/") 'forward slash
   A=Replace(A,"\b","") 'backspace
   A=Replace(A,"\f","") 'form feed
   A=Replace(A,"\n","") 'new line
   A=Replace(A,"\r","") 'carraige return
   A=Replace(A,"\t","") 'tab
   Return A
End Function

Public Sub JSON_Polygon2Rectangle(JS As Object, Key As String, Rectangle As Object, Units As String, XRes As Long, YRes As Long)
   'Microsoft returns the coordinates of a region as JSON ->   "polygon": [1848,492,1896,494,1897,535,1849,535]
   'We need to convert this to  .left, .width, .top and .height
   Rectangle.Left=  min(CDouble(JS(Key & ".boundingBox(0)")),CDouble(JS(Key & ".boundingBox(6)")))
   Rectangle.Width= max(CDouble(JS(Key & ".boundingBox(2)")),CDouble(JS(Key & ".boundingBox(4)")))-Rectangle.Left
   Rectangle.Top =  min(CDouble(JS(Key & ".boundingBox(1)")),CDouble(JS(Key & ".boundingBox(3)")))
   Rectangle.Height=max(CDouble(JS(Key & ".boundingBox(5)")),CDouble(JS(Key & ".boundingBox(7)")))-Rectangle.Top
   If Units="inch" Then
      Rectangle.Left=Rectangle.Left*XRes
      Rectangle.Width=Rectangle.Width*XRes
      Rectangle.Top=Rectangle.Top*YRes
      Rectangle.Height=Rectangle.Height*YRes
   End If
End Sub

Function CDouble(t As String) As Double
   'Convert a string to a double amount safely using the default amount formatter, where you control the decimal separator.
   'Make sure your amount formatter your choose has "." as the decimal symbol as Microsoft OCR returns coordinates in this format: "137.0"
   'CLng and CDbl functions use local regional settings
   Dim F As New CscXDocField, AF As ICscFieldFormatter
   F.Text=t
   Set AF=Project.FieldFormatters.ItemByName("USAmountFormatter")
   AF.FormatField(F)
   Return F.DoubleValue
End Function
