'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
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
   If DefaultPageProfileName="Microsoft OCR" Then Microsoft_FormRecogniser(pXDoc)
End Sub

Public Sub Microsoft_FormRecogniser(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, OCR As String, P As Long, RepName As String, L As String, StartTime As Long
   RepName="MicrosoftOCR"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftComputerVisionKey").Value   'Key to use Microsoft Cognitive Services
   Model=Project.ScriptVariables.ItemByName("MicrosoftFormRecognizerModel").Value
   If pXDoc.CDoc.Pages.Count=pXDoc.CDoc.SourceFiles.Count Then 'if the document has only single page files
         For P=0 To pXDoc.CDoc.Pages.Count-1
         OCR=""
         Open "c:\temp\ocr.json" For Input As #1
         While Not EOF 1
            Line Input #1, L
            OCR = OCR & L
         Wend
         Close #1
         StartTime=Timer
         OCR=MicrosoftFormRecogniser_REST(pXDoc.CDoc.Pages(P).SourceFileName,Model,EndPoint,Key,10)
         'Store time in seconds that Microsoft took to read document
         If pXDoc.XValues.ItemExists("MicrosoftOCR_Time") Then pXDoc.XValues.Delete("MicrosoftOCR_Time")
         pXDoc.XValues.Add("MicrosoftOCR_Time",CStr(Timer-StartTime),True)
         Open "c:\temp\ocr.json" For Output As #1
         Print #1, vbUTF8BOM & OCR
         Close #1
         pXDoc.Representations.Create(RepName)
         MicrosoftOCR_AddWords(pXDoc, OCR, P)
      Next
      Exit Sub
   End If
   For P=0 To pXDoc.CDoc.Pages.Count-1
      Dim img As CscImage
      Dim imgBW As CscImage
      Set img = pXDoc.CDoc.Pages(P).GetImage()
     ' Set imgBW = img.BinarizeWithVRS
      Dim ofs As New FileSystemObject
      Dim tempFile As String
      tempFile = ofs.GetParentFolderName(pXDoc.CDoc.Pages(P).SourceFileName) & "\" _
         & ofs.GetBaseName(pXDoc.CDoc.Pages(P).SourceFileName) & CStr(P) & ".tif" ' & ofs.GetExtensionName(pXDoc.CDoc.Pages(P).SourceFileName)
      img.Save(tempFile,CscImgFileFormatPNG)
      OCR=MicrosoftFormRecogniser_REST(tempFile,Model, EndPoint, Key,10)
      MicrosoftOCR_AddWords(pXDoc, OCR, P)
      Set img = Nothing
      'Kill tempFile
   Next
End Sub

Public Function MicrosoftFormRecogniser_REST(ImageFileName As String, Model As String, EndPoint As String, Key As String,Retries As Long) As String
   'Call Microsoft Azure Form Recognizer 3.0
   'https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/how-to-guides/use-sdk-rest-api?view=doc-intel-3.1.0&tabs=windows&pivots=programming-language-rest-api
   'model = prebuilt-document
   Dim HTTP As New MSXML2.XMLHTTP60, Image() As Byte, I As Long, Delay As Long, RegexAzureStatus As New RegExp, getRequestStatus As MatchCollection, OperationLocation As String
   RegexAzureStatus.Pattern = """status"":""(.*?)"""
   'supports PDF, JPEG, JPG, PNG, TIFF, BMP, 50x50 up to 4200x4200, max 10 megapixel
   Open ImageFileName For Binary Access Read As #1
      ReDim Image (0 To LOF(1)-1)
      Get #1,, Image
   Close #1
   'version=2023-07-31, version=2022-08-31
   HTTP.Open("POST", EndPoint & "formrecognizer/documentModels/" & Model & ":analyze?api-version=2023-07-31&stringIndexType=textElements",varAsync:=False)
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
   For I= 1 To Retries
      Wait Delay
      HTTP.Open("GET", OperationLocation,varAsync:=False)
      HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
      HTTP.send()
      If HTTP.status<>200 Then Err.Raise (655,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
      Select Case getRequestStatus(0).SubMatches(0)
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

Public Sub MicrosoftOCR_AddWords(pXDoc As CscXDocument, OCR As String, PageOffset As Long)
   'Microsoft OCR returns results in this format
   '       {"content":"London","boundingBox":[1577.0,403.0,1643.0,404.0,1641.0,454.0,1575.0,453.0],"confidence":0.988,"span":{"offset":17,"length":4}}
   Dim RegexPages As New RegExp, RegexWords As New RegExp, Confidences As String
   Dim RegexLines As New RegExp
   Dim Pages As MatchCollection, P As Long, PageIndex As Long
   Dim Words As MatchCollection, W As Long, BoundingBox() As String, Confidence As Double, Word As CscXDocWord
   'RegexPages.Pattern="""pageNumber"":(\d+),""words"":\[({.*?})\],""spans"""   'returns pagenumber and words from JSON
   RegexPages.Pattern="""pageNumber"":(\d+),.*?""words"":\[(.*?)\],""lines"""   'returns pagenumber and lines from JSON
   RegexPages.Global=True ' Find more than one page!
   'RegexWords.Pattern="""content"":""(.*?)"",""boundingBox"":\[(.*?)\],""confidence"":(.*?),""span"""  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Pattern="""content"":""(.*?)"",""polygon"":\[(.*?)\],""confidence"":(\d\.\d+),"  ' returns each word, boundingbox coordinates, confidence from a page
   RegexWords.Global=True ' find more than one word!
   Set Pages=RegexPages.Execute(OCR)
   For P=0 To Pages.Count-1
      PageIndex=CLng(Pages(P).SubMatches(0))-1 ' if a page is missing OCR it is possibe that PageNr is not the same as P.
      Set Words = RegexWords.Execute(Pages(P).SubMatches(1))
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
