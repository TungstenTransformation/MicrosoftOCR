'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
Option Explicit
' Project Script
'#Language "WWB-COM"

'This project only supports Microsoft Document Intelligence v4.0 (preview)
'It does not support version 3.1
'https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/how-to-guides/use-sdk-rest-api?view=doc-intel-4.0.0&preserve-view=true&tabs=windows&pivots=programming-language-rest-api

Private Sub Document_BeforeExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   'If we are in TotalAgility ExtractionGroup Project and there is no OCR then force OCR to happen.
   Dim bSkip As Boolean
   If pXDoc.Representations.Count=0 Then Document_BeforeClassifyXDoc(pXDoc, bSkip)
End Sub

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Tungsten Transformation, rename the default page OCR profile to "Microsoft DI"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="Microsoft DI" Then MicrosoftDI(pXDoc)
End Sub

Private Sub Document_AfterClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument)
  'If Microsoft Document Intelligence classified the document, then ignore Transformation Classification and use Microsoft's classification
   Dim Model As String, JSONs As String, JSON As Object, className As String, confidence As Double
   JSONs=Cache_Load(pXDoc,"MicrosoftDI_JSON",False) 'Get the Microsoft DI response JSON if it is there.
   Set JSON=JSON_Parse(JSONs)
   className=JSON("analyzeResult")("documents")(0)("docType")
   confidence=JSON("analyzeResult")("documents")(0)("confidence")
   pXDoc.Reclassify(className,confidence)
End Sub

Public Sub MicrosoftDI(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, RepName As String, StartTime As Long, Cache As String, JSON As String, Model As String, JS As Object, Version As String
   Dim TimeStart As Double, TimeEnd As Double, FileName As String
   RepName="MicrosoftDI"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Microsoft Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceKey").Value   'Key to use Microsoft Cognitive Services
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value 'https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/choose-model-feature?view=doc-intel-4.0.0
   Version=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceAPIVersion").Value  ' this project only supports "2024-02-29-preview"
   'JSON=Cache_Load(pXDoc,"MicrosoftDI_JSON")
   If JSON="" Then
      If pXDoc.CDoc.SourceFiles.Count=1 Then 'Does the XDoc contain 1 or more image files.
         FileName=pXDoc.CDoc.SourceFiles(0).FileName
      Else
         FileName = XDocument_ConvertToMultipageTIFF(pXDoc,False) 'We can send only one image to Microsoft. If we have multiple images,we merge them into a multipage TIFF.
      End If
      StartTime=Timer
      JSON=MicrosoftDI_REST(FileName,Model,EndPoint,Version,Key,10)
      TimeEnd=Timer
      If pXDoc.CDoc.SourceFiles.Count>1 Then Kill FileName 'delete temp multipage tiff
      If TimeEnd<TimeStart Then 'Store time in milliseconds that Microsoft took to read document
         pXDoc.TimeOCR = CLng(1000 * (86400 - TimeStart + TimeEnd)) ' 86400=24*60^2 = seconds/day. needed if the job started before midnight and finished after midnight
      Else
         pXDoc.TimeOCR = CLng(1000 * (TimeEnd - TimeStart))  ' this is in milliseconds (accuracy of 1/18th of a second)
      End If

      If pXDoc.XValues.ItemExists("MicrosoftDI_Time") Then pXDoc.XValues.Delete("MicrosoftDI_Time")
      pXDoc.XValues.Add("MicrosoftDI_Time",CStr(Timer-StartTime),True)
   End If
   Set JS= JSON_Parse(JSON)
   pXDoc.Representations.Create(RepName)
   MicrosoftDI_AddWords(pXDoc, JS, 0)
   Cache_Save(pXDoc,"MicrosoftDI_JSON",JSON)
End Sub

Public Function MicrosoftDI_REST(ImageFileName As String, Model As String, EndPoint As String, Version As String, Key As String,Retries As Long) As String
   'Call Microsoft Azure Form Recognizer 3.0
   'https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/how-to-guides/use-sdk-rest-api?view=doc-intel-3.1.0&tabs=windows&pivots=programming-language-rest-api
   'model = prebuilt-document
   Dim HTTP As MSXML2.XMLHTTP60, Image() As Byte, I As Long, Delay As Long, RegexAzureStatus As New RegExp, getRequestStatus As MatchCollection, OperationLocation As String, status As String, URL As String, Extension As String
   RegexAzureStatus.Pattern = """(?:message|status)"":\s*""(.*?)""" 'Get message or status from JSON via regex
   'supports PDF, JPEG, JPG, PNG, TIFF, BMP, 50x50 up to 4200x4200, max 10 megapixel
   Open ImageFileName For Binary Access Read As #1
      ReDim Image (0 To LOF(1)-1)
      Get #1,, Image
   Close #1
   URL=EndPoint & "/documentintelligence/documentModels/" & Model & ":analyze?_overload=analyzeDocument&api-version=" & Version
   Extension =LCase(Mid(ImageFileName,InStrRev(ImageFileName,".")+1))
   If Extension="jpg" Then Extension = "jpeg"
   If Extension="tif" Then Extension = "tiff"
   Select Case Extension
   Case "pdf"
         Extension="application/pdf"
   Case "png", "jpeg", "tiff", "bmp"
         Extension="image/" & Extension
   Case Else
         'Extension="application/octet-stream"
         Err.Raise(658,"Unsupported file type " & Extension)
   End Select
   Delay=1 'Wait 1 second for result (Microsoft recommends calling no more frequently than 1 second)
   For I = 1 To 100
      Set HTTP = New MSXML2.XMLHTTP60
      HTTP.Open("POST", URL ,varAsync:=False)
      HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
      HTTP.setRequestHeader("Content-Type", Extension)
      HTTP.send(Image)
      Debug.Print HTTP.status
      Select Case HTTP.status
         Case 202 'success
            Exit For
         Case 429 'exceeded call rate per second
            Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
            Wait 3 * Delay 'seconds and try again
         Case Else
            Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText)
            Err.Raise (654,,"Microsoft OCR Error: (" & HTTP.status & ") " & getRequestStatus(0).SubMatches(0))
      End Select
   Next
   If I>90 Then Err.Raise (657,,"Microsoft OCR Error: (" & HTTP.status & ") " & getRequestStatus(0).SubMatches(0))
   OperationLocation=HTTP.getResponseHeader("Operation-Location") 'Get the URL To retrive the result
   For I= 1 To 100
      Wait Delay
      Set HTTP = New MSXML2.XMLHTTP60
      HTTP.Open("GET", OperationLocation  & "&a=" & CStr(I) ,varAsync:=False) 'pass a random parameter each time so Windows doesn't cache
      HTTP.setRequestHeader("Ocp-Apim-Subscription-Key", Key)
      HTTP.send()
      Set getRequestStatus = RegexAzureStatus.Execute(HTTP.responseText) ' find the status or message in json response.
      status=getRequestStatus(0).SubMatches(0)
      Select Case status
         Case "succeeded" '200
            Exit For
         Case "running", "notStarted" 'also 200
            'do nothing as job not finished
            Delay=1
         Case Else 'error
            If HTTP.status=429 Then 'to many requests at this license level
               Wait 3 * Delay ' seconds
            Else
               Err.Raise (655,,"Microsoft OCR Error: (" & HTTP.status & ") " & HTTP.responseText)
            End If
      End Select
   Next
   Return HTTP.responseText
End Function

Public Sub MicrosoftDI_AddWords(pXDoc As CscXDocument, JS As Object, PageOffset As Long, Optional UseMicrosoftTextLines As Boolean)
   Dim p As Long, W As Long, Confidences As String, Word As CscXDocWord, Units As String, XRes As Double, YRes As Double
   Dim pages As Object, ocrWord As Object
   Set pages=JS("analyzeResult")("pages")
   For p=0 To pages.Count-1
      Units=pages(p)("unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(p).XRes
         YRes=pXDoc.CDoc.Pages(p).XRes
      End If
      For W=0 To pages(p)("words").Count-1   'format
         Set ocrWord = pages(p)("words")(W)
          Set Word = New CscXDocWord
          Word.Text=ocrWord("content")
          Word.PageIndex=p
          BoundingBox2Rectangle(ocrWord("polygon"),Word,Units,XRes,YRes) 'Give the words the correct coordinates
          Confidences = Confidences & Format("0.000", ocrWord("confidence")) & ","
          pXDoc.Pages(p+PageOffset).AddWord(Word)
      Next 'Word
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
      Set ocrWord = pages(Word.PageIndex)("lines")(Word.LineIndex)("words")(Word.IndexInTextLine)
      Units=pages(Word.PageIndex)("unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(p).XRes
         YRes=pXDoc.CDoc.Pages(p).XRes
      End If
      BoundingBox2Rectangle(ocrWord("boundingBox"),Word,Units,XRes,YRes)
   Next
End Sub

Public Function Min(A,b)
   'test
   If A<b Then Min=A Else Min=b
End Function
Public Function Max(A,b)
   If A>b Then Max=A Else Max=b
End Function

Private Function XDocument_ConvertToMultipageTIFF(ByVal pXDoc As CASCADELib.CscXDocument, Optional ByVal Bitonal As Boolean=True)
   'Microsoft Document Intelligence API only supports sending a single document.
   'See "Request Body" at https://westus.dev.cognitive.microsoft.com/docs/services/form-recognizer-api-2023-07-31/operations/AnalyzeDocument
   'It supports pdf, jpeg, png, tiff, bmp, text, docx, xlsx, pptx, but not multiple images. If we have singlepage images in one document, we need to merge to multipage tiff.
   Dim NewImg As New CscImage, SourceImg As CscImage, TargetImgPath As String
   Dim p As Long, FileFormat As Long, ColorFormat As Long

   For p = 0 To pXDoc.CDoc.Pages.Count - 1
      'Derive new filename from existing name - just replace extension with .tif
      With pXDoc.CDoc.Pages(p)
         Set SourceImg=pXDoc.CDoc.Pages(.IndexOfSourceFile).GetImage()
         If p = 0 Then
            TargetImgPath = Left(.SourceFileName,InStrRev(.SourceFileName,"\")) & "multipage.tif"
            If Bitonal Then
               FileFormat=CscImageFileFormat.CscImgFileFormatTIFFFaxG4
               ColorFormat=CscImageColorFormat.CscImgColFormatBinary
            Else
               Select Case SourceImg.BitsPerSample
               Case 1
                  FileFormat=CscImageFileFormat.CscImgFileFormatTIFFFaxG4
                  ColorFormat=CscImageColorFormat.CscImgColFormatBinary
               Case 8
                  FileFormat=CscImageFileFormat.CscImgFileFormatTIFFUncompressed
                  ColorFormat=CscImageColorFormat.CscImgColFormatGray8
               Case 16
                  FileFormat=CscImageFileFormat.CscImgFileFormatTIFFUncompressed
                  ColorFormat=CscImageColorFormat.CscImgColFormatGray16
               Case 24
                  FileFormat=CscImageFileFormat.CscImgFileFormatTIFFUncompressed
                  ColorFormat=CscImageColorFormat.CscImgColFormatRGB24
               End Select
            End If
           'for the first image, mark the new tiff to remain open for new pages to be added
           NewImg.StgFilterControl(FileFormat, CscStgCtrlTIFFKeepFileOpen, TargetImgPath, 0, 0)
         End If
      End With
      ' load current page image into the new image that is kept open
      If Bitonal Then Set SourceImg=SourceImg.BinarizeWithVRS()
      NewImg.CreateImage(ColorFormat, SourceImg.Width, SourceImg.Height, SourceImg.XResolution, SourceImg.YResolution)
      NewImg.CopyRect(SourceImg, 0, 0, 0, 0, SourceImg.Width, SourceImg.Height)
      ' save new image (as KeepFileOpen was set, this will append to the existing file)
      NewImg.Save(TargetImgPath, FileFormat)
   Next
   'close the multi-page TIFF
   NewImg.StgFilterControl(FileFormat, CscStgCtrlTIFFCloseFile, TargetImgPath, 0, 0)
   Return TargetImgPath
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

Public Function Cache_Load(pXDoc As CscXDocument, RepName As String,Optional Retrieve As Boolean=True) As String
   'Return the Microsoft DI JSON result from the Xdoc Representation it was stored in.
   'if retrieve = false then DON'T download it again from Microsoft DI
   Dim R As Long, Model As String, CacheFileName As String
   For R=0 To pXDoc.Representations.Count-1
      If pXDoc.Representations(R).Name=RepName Then Return pXDoc.Representations(R).Words(0).Text
   Next
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value
   CacheFileName=Replace(pXDoc.FileName,".xdc", "." & Model & ".json")
   If Dir(CacheFileName)<>"" And Retrieve Then Return File_Load(CacheFileName)
End Function

Public Sub Cache_Save(pXDoc As CscXDocument, RepName As String, Content As String)
   'Cache a string
   Dim R As Long, Model As String, CacheFileName As String
   For R=pXDoc.Representations.Count-1 To 0 Step -1
      If pXDoc.Representations(R).Name=RepName Then pXDoc.Representations.Remove(R)
   Next
   With pXDoc.Representations.Create(RepName).Words
      Dim Word As New CscXDocWord
      Word.Text=Content
      .Append(Word)
   End With
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value
   CacheFileName=Replace(pXDoc.FileName,".xdc", "." & Model & ".json")
   Open CacheFileName For Output As #1
   Print #1, vbUTF8BOM & Content
   Close #1
End Sub

Public Sub BoundingBox2Rectangle(bb As Object, Rectangle As Object, Units As String, XRes As Long, YRes As Long)
   'Microsoft returns the coordinates of a region as JSON ->   "polygon": [1848,492,1896,494,1897,535,1849,535]
   'We need to convert this to  .left, .width, .top and .height
   Dim L As Double, W As Double, T As Double, H As Double
   L= Min(bb(0),bb(6))
   W= Max(bb(2),bb(4))-L
   T= Min(bb(1),bb(3))
   H= Max(bb(5),bb(7))-T
   If Units="inch" Then
      Rectangle.Left=L*XRes
      Rectangle.Width=W*XRes
      Rectangle.Top=T*YRes
      Rectangle.Height=H*YRes
   Else
      Rectangle.Left=L
      Rectangle.Width=W
      Rectangle.Top=T
      Rectangle.Height=H
   End If
End Sub


'-------------------------------------------------------------------
' VBA JSON Parser https://github.com/KofaxTransformation/KTScripts/blob/master/JSON%20parser%20in%20vb.md
'-------------------------------------------------------------------

'This converts a JSON string into a hierarchy of Dictionary, SortedList, String, Double, True, False and Nothing objects that are easy to navigate and loop through.
Private T As Long, Tokens As Object ' The JSON is broken into an array of tokens. T is the current index of the parser
Function JSON_Parse(JSON As String, Optional Key As String = "$") As Object
   'This is 100% compliant with ECMA-404 JSON Data Interchange Standard at https://www.json.org/json-en.html
   'the regex pattern finds strings including characters escaped with \ OR numbers OR true/false/null OR \\{}:,[]
   'tested at https://regex101.com/r/YkiVdc/1
   'This script will crash on invalid JSON
   With CreateObject("vbscript.regexp")
      .Global=True
      'This regex completely splits any JSON into an array of tokens - a token is any of these 6 characters {}[]:, or string/number/true/false/null.
      'The order of sections in the regex ensures that it parses correctly because escaped characters are parsed first.

      '   JSON =        String        OR               Number               OR  true/false/null OR  []{}:,
      .Pattern = """(?:[^""\\]|\\.)*""|-?(?:\d+)(?:\.\d*)?(?:[eE][+\-]?\d+)?|(?:true|false|null)|[\[\]{}:,]"
      'The ?: means "non-capturing group". This gives then a 1-dimension array of matching subgroups, instead of a 2-dimensional array of groups and subgroups.
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
   If tokens(T+1)="}" Then  T=T+2 : Return Obj ' empty object
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
