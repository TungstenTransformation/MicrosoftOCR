Option Explicit
   'On Menu/Edit/References... Add reference to "Microsoft XML, v 6.0"
   'On Menu/Edit/References... Add reference to "Microsoft VB Regular Expressions 5.5"

' Project Script

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Kofax Transformation, rename the default page OCR profile to "Microsoft OCR"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="Microsoft OCR" Then MicrosoftOCR_Read(pXDoc)
End Sub

Public Sub MicrosoftOCR_Read(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, OCR As String, P As Long, RepName as String
   RepName="Microsoft OCR"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   pXDoc.Representations.Create(RepName)
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
   For P=0 To pXDoc.CDoc.Pages.Count-1
      OCR=MicrosoftOCR_REST(pXDoc.CDoc.Pages(P).SourceFileName,EndPoint,Key)
      MicrosoftOCR_AddWords(pXDoc, OCR, P)
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
   HTTP.Open("POST", EndPoint & "/computervision/imageanalysis:analyze?features=Read&model-version=latest&api-version=2022-10-12-preview",varAsync:=False)
   HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
   HTTP.setRequestHeader("Content-Type", "application/octet-stream")
   HTTP.send(Image)
   If HTTP.status<>200 Then Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
   MicrosoftOCR_REST = HTTP.responseText
End Function

Public Sub MicrosoftOCR_AddWords(pXDoc As CscXDocument, OCR As String, PageOffset As Long)
   Dim RegexPages As New RegExp, RegexWords As New RegExp
   Dim Pages As MatchCollection, P As Long, PageIndex As Long
   Dim Words As MatchCollection, W As Long, BoundingBox() As String, Confidence As Double, Word As CscXDocWord
   RegexPages.Pattern="""pageNumber"":(\d+),""words"":\[({.*?})\],""spans"""   'returns pagenumber and words from JSON
   RegexPages.Global=True ' Find more than one page!
   RegexWords.Pattern="""content"":""(.*?)"",""boundingBox"":\[(.*?)\],""confidence"":(.*?),""span"""  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Global=True ' find more than one word!
   Set Pages=RegexPages.Execute(OCR)
   For P=0 To Pages.Count-1
      PageIndex=CLng(Pages(P).SubMatches(0))-1 ' if a page is missing OCR it is possibe that PageNr is not the same as P.
      Set Words = RegexWords.Execute(Pages(P).SubMatches(1))' The JSON of all the words and coordinates on this page
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
         'Word.Confidence = 1.0' CDouble(Words(W).SubMatches(2))
         pXDoc.Pages(PageIndex+PageOffset).AddWord(Word)
      Next
   Next
   pXDoc.Representations(0).AnalyzeLines 'Redo Text Line Analysis in Kofax Transformation
End Sub

Public Function min(a,b)
   If a<b Then min=a Else min=b
End Function
Public Function max(a,b)
   If a>b Then max=a Else max=b
End Function

Function CDouble(t As String) As Double
   'Convert a string to a double amount safely using the default amount formatter, where you control the decimal separator.
   'CLng and CDbl functions use local regional settings
   Dim F As New CscXDocField
   F.Text=t
   DefaultAmountFormatter.FormatField(F)
   Return F.DoubleValue
End Function

Public Function JSON_Unescape(a As String) As String
   'https://www.json.org/json-en.html
   a=Replace(a,"\""","""") 'double quote
   a=Replace(a,"\\","\") 'backslash
   a=Replace(a,"\/","/") 'forward slash
   a=Replace(a,"\b","") 'backspace
   a=Replace(a,"\f","") 'form feed
   a=Replace(a,"\n","") 'new line
   a=Replace(a,"\r","") 'carraige return
   a=Replace(a,"\t","") 'tab
   JSON_Unescape=a
End Function
