'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
Option Explicit
   'On Menu/Edit/References... Add reference to "Microsoft XML, v 6.0"
   'On Menu/Edit/References... Add reference to "Microsoft VB Regular Expressions 5.5"

' Project Script
'#Language "WWB-COM"

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Kofax Transformation, rename the default page OCR profile to "Microsoft OCR"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="Microsoft OCR" Then Microsoft_FormRecogniser(pXDoc)
End Sub

Public Sub Microsoft_FormRecogniser(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, P As Long, RepName As String, StartTime As Long, Cache As String, JSON As String, Model As String
   RepName="MicrosoftOCR"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceKey").Value   'Key to use Microsoft Cognitive Services
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value
   Cache=Replace(pXDoc.FileName,".xdc", "." & Model & ".json")
   If pXDoc.CDoc.Pages.Count=pXDoc.CDoc.SourceFiles.Count Then 'if the document has only single page files
      For P=0 To pXDoc.CDoc.Pages.Count-1
         If Dir(Cache)<>"" Then
            JSON=File_Load(Cache)
         Else
            StartTime=Timer
            JSON=MicrosoftFormRecogniser_REST(pXDoc.CDoc.Pages(P).SourceFileName,Model,EndPoint,Key,10)
            'Store time in seconds that Microsoft took to read document
            If pXDoc.XValues.ItemExists("MicrosoftOCR_Time") Then pXDoc.XValues.Delete("MicrosoftOCR_Time")
            pXDoc.XValues.Add("MicrosoftOCR_Time",CStr(Timer-StartTime),True)
            Open Cache For Output As #1
            Print #1, vbUTF8BOM & JSON
            Close #1
         End If
         Dim JS As Object
         Set JS= JSON_Parse(JSON)
         pXDoc.Representations.Create(RepName)
         MicrosoftOCR_AddWords(pXDoc, JS, P)
      Next
      Exit Sub
   End If
End Sub

Public Function MicrosoftFormRecogniser_REST(ImageFileName As String, Model As String, EndPoint As String, Key As String,Retries As Long) As String
   'Call Microsoft Azure Form Recognizer 3.0
   'https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/how-to-guides/use-sdk-rest-api?view=doc-intel-3.1.0&tabs=windows&pivots=programming-language-rest-api
   'model = prebuilt-document
   Dim HTTP As New MSXML2.XMLHTTP60, Image() As Byte, I As Long, Delay As Long, RegexAzureStatus As New RegExp, getRequestStatus As MatchCollection, OperationLocation As String, status As String, URL As String
   RegexAzureStatus.Pattern = """status"":""(.*?)"""
   'supports PDF, JPEG, JPG, PNG, TIFF, BMP, 50x50 up to 4200x4200, max 10 megapixel
   Open ImageFileName For Binary Access Read As #1
      ReDim Image (0 To LOF(1)-1)
      Get #1,, Image
   Close #1
   'version=2023-07-31, version=2022-08-31
   URL=EndPoint & "formrecognizer/documentModels/" & Model & ":analyze?api-version=2023-07-31&stringIndexType=textElements"
   HTTP.Open("POST", URL ,varAsync:=False)
   HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
   If LCase(Right(ImageFileName,3))="pdf" Then
         HTTP.setRequestHeader("Content-Type", "application/pdf")
      Else
         HTTP.setRequestHeader("Content-Type", "application/octet-stream")
   End If
   HTTP.send(Image)
   If HTTP.status<>202 Then
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
      Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & getRequestStatus(0).SubMatches(0))
   End If
   OperationLocation=HTTP.getResponseHeader("Operation-Location") 'Get the URL To retrive the result
   Delay=1 'Wait 1 second for result (Microsoft recommends calling no more frequently than 1 second)
   For I= 1 To 100
      Wait Delay
      Set HTTP = New MSXML2.XMLHTTP60
      HTTP.Open("GET", OperationLocation,varAsync:=False)
      HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
      HTTP.send()
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
      status=getRequestStatus(0).SubMatches(0)
      If HTTP.status<>200 Then Err.Raise (655,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
      Select Case status
      Case "succeeded"
         Exit For
      Case "failed"
            Err.Raise (656,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
      Case "running", "notStarted"
         Delay=Delay+1 ' wait 1 second longer next time
      End Select
   Next

   Return HTTP.responseText
End Function

Public Sub MicrosoftOCR_AddWords(pXDoc As CscXDocument, JS As Object, PageOffset As Long)
   Dim P As Long, W As Long, Key As String, Confidences As String, Word As CscXDocWord
   For P=0 To JS("js.analyzeResult.pages._count")-1
      For W=0 To JS("js.analyzeResult.pages(" & P & ").words._count")-1   'format
         Key="js.analyzeResult.pages(" & P & ").words(" & W & ")"
         Set Word = New CscXDocWord
         Word.Text=JSON_Unescape(JS(Key & ".content"))
         Word.PageIndex=P
         Word.Left=  min(CDouble(JS(Key & ".polygon(0)")),CDouble(JS(Key & ".polygon(6)")))
         Word.Width= max(CDouble(JS(Key & ".polygon(2)")),CDouble(JS(Key & ".polygon(4)")))-Word.Left
         Word.Top =  min(CDouble(JS(Key & ".polygon(1)")),CDouble(JS(Key & ".polygon(3)")))
         Word.Height=max(CDouble(JS(Key & ".polygon(5)")),CDouble(JS(Key & ".polygon(7)")))-Word.Top
         Confidences = Confidences & JS(Key & ".confidence") & ","
         pXDoc.Pages(P+PageOffset).AddWord(Word)
      Next
   Next
   Confidences = Left(Confidences,Len(Confidences)-1) ' trim trailing ,
   'Store all confidences for later use in AZL
   If pXDoc.XValues.ItemExists("MicrosoftOCR_WordConfidences") Then pXDoc.XValues.Delete("MicrosoftOCR_WordConfidences")
   pXDoc.XValues.Add("MicrosoftOCR_WordConfidences",Confidences,True)
   pXDoc.Representations(0).AnalyzeLines 'Redo Text Line Analysis in Kofax Transformation
End Sub

Public Sub MicrosoftOCR_AddTables(pXDoc As CscXDocument, JS As Object, PageOffset As Long)

End Sub

Public Sub MicrosoftOCR_AddTable(pXDoc As CscXDocument, JS As Object, Table As CscXDocTable, T As Long)
   Dim Row As CscXDocTableRow, R As Long, C As Long, CellIndex As Long, Cell As CscXDocTableCell, W As Long, Words As CscXDocWords, P As Long, Key As String, BR As Long, BRKey As String
   Dim rowCount As Long, columnCount As Long
   Table.Rows.Clear
   rowCount =CLng(JS("js.analyzeResult.tables(" & T & ").rowCount"))
   While Table.Rows.Count<rowCount
      Table.Rows.Append
   Wend
   columnCount = CLng(JS("js.analyzeResult.tables(" & T & ").columnCount"))
   For CellIndex =0 To rowCount*columnCount-1
      Key="js.analyzeResult.tables(" & T & ").cells(" & CellIndex & ")"
      R=CLng(JS(Key & ".rowIndex"))
      C=CLng(JS(Key & ".columnIndex"))
      Set Cell=Table.Rows(R).Cells(C)
      'Cell.Text=JSON_Unescape(JS(Key & ".content"))
      For BR = 0 To CLng(JS(Key & ".boundingRegions._count"))-1
         BRKey = Key & ".boundingRegions(" & BR & ")"
         P =CLng(JS(BRKey & ".pageNumber"))-1
         Cell.Left=  min(CDouble(JS(BRKey & ".polygon(0)")),CDouble(JS(BRKey & ".polygon(6)")))
         Cell.Width= max(CDouble(JS(BRKey & ".polygon(2)")),CDouble(JS(BRKey & ".polygon(4)")))-Cell.Left
         Cell.Top =  min(CDouble(JS(BRKey & ".polygon(1)")),CDouble(JS(BRKey & ".polygon(3)")))
         Cell.Height=max(CDouble(JS(BRKey & ".polygon(5)")),CDouble(JS(BRKey & ".polygon(7)")))-Cell.Top
         Set Words = pXDoc.GetWordsInRect(P,Cell.Left,Cell.Top, Cell.Width, Cell.Height)
         For W=0 To Words.Count-1
            Cell.AddWordData(Words(W))
         Next
      Next
   Next
End Sub

Public Sub MicrosoftOCR_AddWords2(pXDoc As CscXDocument, JSON As String, PageOffset As Long)
   'Microsoft OCR returns results in this format
   '       {"content":"London","boundingBox":[1577.0,403.0,1643.0,404.0,1641.0,454.0,1575.0,453.0],"confidence":0.988,"span":{"offset":17,"length":4}}
   Dim RegexPages As New RegExp, RegexWords As New RegExp, Confidences As String
   Dim RegexLines As New RegExp
   Dim pages As MatchCollection, P As Long, PageIndex As Long
   Dim Words As MatchCollection, W As Long, BoundingBox() As String, Confidence As Double, Word As CscXDocWord
   'RegexPages.Pattern="""pageNumber"":(\d+),""words"":\[({.*?})\],""spans"""   'returns pagenumber and words from JSON
   RegexPages.Pattern="""pageNumber"":(\d+),.*?""words"":\[(.*?)\],""lines"""   'returns pagenumber and lines from JSON
   RegexPages.Global=True ' Find more than one page!
   'RegexWords.Pattern="""content"":""(.*?)"",""boundingBox"":\[(.*?)\],""confidence"":(.*?),""span"""  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Pattern="""content"":""(.*?)"",""polygon"":\[(.*?)\],""confidence"":(\d\.\d+),"  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Global=True ' find more than one word!
   Set pages=RegexPages.Execute(JSON)
   For P=0 To pages.Count-1
      PageIndex=CLng(pages(P).SubMatches(0))-1 ' if a page is missing OCR it is possibe that PageNr is not the same as P.
      Set Words = RegexWords.Execute(pages(P).SubMatches(1))
      For W=0 To Words.Count-1 ' Create a Kofax Transformation word for each OCR word
         Set Word = New CscXDocWord
         Word.Text=JSON_Unescape( Words.Item(W).SubMatches(0))
         Word.PageIndex=PageIndex+PageOffset
         BoundingBox=Split(Words(W).SubMatches(1),",")' returns 8 numbers= 4 coordinates of topleft, topright, bottomright and bottomleft of word in pixels
         'Microsoft OCR returns an irregular 4-edged polygon. Kofax Transformation requires a rectangle
         Word.Left=min(CDouble(BoundingBox(0)), CDouble(BoundingBox(6)))
         Word.Width=max(CDouble(BoundingBox(2)), CDouble(BoundingBox(4)))-Word.Left
         Word.Top=min(CDouble(BoundingBox(1)),CDouble(BoundingBox(3)))
         Word.Height=max(CDouble(BoundingBox(5)), CDouble(BoundingBox(7)))-Word.Top
         Word.StringTag=Words(W).SubMatches(2)
         ' We cannot set Word.Confidence directly from script, so we will store confidences in an XValue for the AZL
         Confidences = Confidences & Words(W).SubMatches(2) & ","
         pXDoc.Pages(PageIndex+PageOffset).AddWord(Word)
      Next W
   Next P
   Confidences = Left(Confidences,Len(Confidences)-1) ' trim trailing ,
   'Store all confidences for later use in AZL
   If pXDoc.XValues.ItemExists("MicrosoftOCR_WordConfidences") Then pXDoc.XValues.Delete("MicrosoftOCR_WordConfidences")
   pXDoc.XValues.Add("MicrosoftOCR_WordConfidences",Confidences,True)
   pXDoc.Representations(0).AnalyzeLines 'Redo Text Line Analysis in Kofax Transformation
End Sub

Public Function min(a,b)
   'test
   If a<b Then min=a Else min=b
End Function
Public Function max(a,b)
   If a>b Then max=a Else max=b
End Function

Function CDouble(T As String) As Double
   'Convert a string to a double amount safely using the default amount formatter, where you control the decimal separator.
   'Make sure your amount formatter your choose has "." as the decimal symbol as Microsoft OCR returns coordinates in this format: "137.0"
   'CLng and CDbl functions use local regional settings
   Dim F As New CscXDocField, AF As ICscFieldFormatter
   F.Text=T
   Set AF=Project.FieldFormatters.ItemByName("DefaultAmountFormatter")
   AF.FormatField(F)
   Return F.DoubleValue
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
                        t = t + 1
                        dic.Add(Key, "null")
                    Else
                        JSON_ParseObj(Key)
                    End If

         Case "}":  Key = JSON_ParentPath(Key): Exit Do
         Case ":":  Key = Key & "." & tokens(t - 1) 'previous token was a key - remember it
         Case ",":  Key = JSON_ParentPath(Key)
         Case Else 'we are in a string. if next is not ":" then we are value - so add to dict!
            If tokens(t + 1) <> ":" Then dic.Add(Key, tokens(t))
     End Select
    Loop
End Function
Function JSON_ParseArr(Key$)
   Dim A As Long
   Do
      t = t + 1
      Select Case tokens(t)
         Case "}"
         Case "{":  JSON_ParseObj(Key & JSON_ArrayID(A))
         Case "[":  JSON_ParseArr(Key)
         Case "]":  Exit Do
         Case ":":  Key = Key & JSON_ArrayID(A)
         Case ",":  A = A + 1
         Case Else: dic.Add(Key & JSON_ArrayID(A), tokens(t))
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

Public Function File_Load(FileName As String) As String
   Dim L As String
   Open FileName For Input As #1
   While Not EOF 1
      Line Input #1, L
      File_Load = File_Load & L
   Wend
   Close #1
End Function
