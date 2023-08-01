'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
'Option Explicit
   'On Menu/Edit/References... Add reference to "Microsoft XML, v 6.0"
   'On Menu/Edit/References... Add reference to "Microsoft VB Regular Expressions 5.5"

' Project Script
'#Language "WWB-COM"

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Kofax Transformation, rename the default page OCR profile to "Microsoft OCR"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="MicrosoftOCR" Then MicrosoftOCR_Read(pXDoc)
End Sub

Public Sub MicrosoftOCR_Read(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, OCR As String, P As Long, RepName As String
   RepName="MicrosoftOCR"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   pXDoc.Representations.Create(RepName)
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
If Project.ScriptExecutionMode = CscScriptExecutionMode.CscScriptModeServerDesign Then
   For P=0 To pXDoc.CDoc.Pages.Count-1
      Dim img As CscImage
      Dim imgBW As CscImage
      Set img = pXDoc.CDoc.Pages(P).GetImage()
      Set imgBW = img.BinarizeWithVRS
      Dim ofs As New FileSystemObject
      Dim tempFile As String
      tempFile = ofs.GetParentFolderName(pXDoc.CDoc.Pages(P).SourceFileName) & "\" _
         & ofs.GetBaseName(pXDoc.CDoc.Pages(P).SourceFileName) & CStr(P) & ".tif" ' & ofs.GetExtensionName(pXDoc.CDoc.Pages(P).SourceFileName)
      imgBW.Save(tempFile, CscImgFileFormatTIFFFaxG4)
      OCR=MicrosoftOCR_REST(tempFile, EndPoint, Key)
      MicrosoftOCR_AddWords(pXDoc, OCR, P)
      Set img = Nothing
      'Kill tempFile
   Next
Else
   For P=0 To pXDoc.CDoc.Pages.Count-1
      OCR=MicrosoftOCR_REST(pXDoc.CDoc.Pages(P).SourceFileName,EndPoint,Key)
      MicrosoftOCR_AddWords(pXDoc, OCR, P)
   Next
End If
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
   'HTTP.Open("POST", EndPoint & "/computervision/imageanalysis:analyze?features=Read&model-version=latest&api-version=2022-10-12-preview",varAsync:=False)
   HTTP.Open("POST", EndPoint & "/vision/v3.2/read/analyze", varAsync:=False)
   HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
   HTTP.setRequestHeader("Content-Type", "application/octet-stream")
   HTTP.send(Image)
   If HTTP.status = 200 Then
      Return HTTP.responseText
   ElseIf HTTP.status = 202 Then
      'Get the call back from the header in the response
      Dim operationStatus As String
      operationStatus = HTTP.getResponseHeader("Operation-Location")
      Dim retries As Integer
      retries = 10
      For i = 1 To retries
         Wait 1
         HTTP.Open("GET", operationStatus, varAsync:= False)
         HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
         HTTP.setRequestHeader("Content-Type", "application/octet-stream")
         HTTP.send ""
         If HTTP.status <> 200 Then
            Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
         Else
            Dim regexStatus As New RegExp
            Dim statuses As MatchCollection
            Dim status As Match
            Dim stat As SubMatches
            regexStatus.Pattern = "(""status"")(:)("")([A-z]*)("")"
            Set statuses = regexStatus.Execute(HTTP.responseText)
            If statuses.Count > 0 Then
               Set status = statuses.Item(0)
               Set stat = status.SubMatches
               If stat.Item(3) = "succeeded" Then
                  Return HTTP.responseText
               ElseIf stat(3) = "failed" Then
                  Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
               End If
            End If
         End If
         Debug.Print HTTP.responseText
      Next i
      Debug.Print "waited " & CStr(i) & " times"
   Else
      Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
   End If
End Function

Public Sub MicrosoftOCR_AddWords(pXDoc As CscXDocument, OCR As String, PageOffset As Long)
   'Microsoft OCR returns results in this format
   '       {"content":"London","boundingBox":[1577.0,403.0,1643.0,404.0,1641.0,454.0,1575.0,453.0],"confidence":0.988,"span":{"offset":17,"length":4}}
   Dim RegexPages As New RegExp, RegexWords As New RegExp, Confidences As String
   Dim RegexLines As New RegExp
   Dim Pages As MatchCollection, P As Long, PageIndex As Long
   Dim Lines As MatchCollection, L As Long
   Dim Words As MatchCollection, W As Long, BoundingBox() As String, Confidence As Double, Word As CscXDocWord
   'RegexPages.Pattern="""pageNumber"":(\d+),""words"":\[({.*?})\],""spans"""   'returns pagenumber and words from JSON
   RegexPages.Pattern="""page"":(\d+),.*?""lines"":\[(.*?\]}\]})"   'returns pagenumber and lines from JSON
   RegexPages.Global=True ' Find more than one page!
   RegexLines.Pattern= """words"":\[.*?}\]}"  ' returns words on each line, boundingbox coordinates, confidence from a page
   RegexLines.Global = True
   'RegexWords.Pattern="""content"":""(.*?)"",""boundingBox"":\[(.*?)\],""confidence"":(.*?),""span"""  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Pattern= "{""boundingBox"":\[(.*?)\],""text"":""(.*?)"",""confidence"":(.*?)}"  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Global=True ' find more than one word!
   Set Pages=RegexPages.Execute(OCR)
   For P=0 To Pages.Count-1
      PageIndex=CLng(Pages(P).SubMatches(0))-1 ' if a page is missing OCR it is possibe that PageNr is not the same as P.
      Set Lines = RegexLines.Execute(Pages(P).SubMatches(1))
      For L = 0 To Lines.Count - 1
         'Set Words = RegexWords.Execute(Pages(P).SubMatches(1))' The JSON of all the words and coordinates on this page
         Set Words = RegexWords.Execute(Lines(L).Value)' The JSON of all the words and coordinates on this line
         For W=0 To Words.Count-1 ' Create a Kofax Transformation word for each OCR word
            Set Word = New CscXDocWord
            'Word.Text=JSON_Unescape( Words.Item(W).SubMatches(0))
            Word.Text=JSON_Unescape( Words.Item(W).SubMatches(1))
            Word.PageIndex=PageIndex+PageOffset
            'BoundingBox=Split(Words(W).SubMatches(1),",")' returns 8 numbers= 4 coordinates of topleft, topright, bottomright and bottomleft of word in pixels
            BoundingBox=Split(Words(W).SubMatches(0),",")' returns 8 numbers= 4 coordinates of topleft, topright, bottomright and bottomleft of word in pixels
            'Microsoft OCR returns an irregular 4-edged polygon. Kofax Transformation requires a rectangle
            Word.Left=min(CDouble(BoundingBox(0)), CDouble(BoundingBox(6)))
            Word.Width=max(CDouble(BoundingBox(2)), CDouble(BoundingBox(4)))-Word.Left
            Word.Top=min(CDouble(BoundingBox(1)),CDouble(BoundingBox(3)))
            Word.Height=max(CDouble(BoundingBox(5)), CDouble(BoundingBox(7)))-Word.Top
            Word.StringTag=Words(W).SubMatches(2)
            ' We cannot set Word.Confidence directly from script, so we will store confidences in an XValue for the AZL
            Confidences = Confidences & Words(W).SubMatches(2) & ","
            'Word.Confidence = 1.0' CDouble(Words(W).SubMatches(2))
            pXDoc.Pages(PageIndex+PageOffset).AddWord(Word)
         Next W
      Next L
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

Function CDouble(t As String) As Double
   'Convert a string to a double amount safely using the default amount formatter, where you control the decimal separator.
   'Make sure your amount formatter your choose has "." as the decimal symbol as Microsoft OCR returns coordinates in this format: "137.0"
   'CLng and CDbl functions use local regional settings
   Dim F As New CscXDocField, AF As ICscFieldFormatter
   F.Text=t
   Set AF=Project.FieldFormatters.ItemByName("DefaultAmountFormatter")
   AF.FormatField(F)
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
   Return a
End Function
