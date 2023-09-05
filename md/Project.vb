'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\SysWOW64\msxml6.dll#Microsoft XML, v6.0#MSXML2
Option Explicit
' Project Script
'#Language "WWB-COM"

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   'To trigger Microsoft Azure OCR in Kofax Transformation, rename the default page OCR profile to "Microsoft DI"
   Dim DefaultPageProfileName As String
   DefaultPageProfileName=Project.RecogProfiles.ItemByID(Project.RecogProfiles.DefaultProfileIdPr).Name
   If DefaultPageProfileName="Microsoft DI" Then MicrosoftDI(pXDoc)
End Sub

Public Sub MicrosoftDI(pXDoc As CscXDocument)
   Dim EndPoint As String, Key As String, RepName As String, StartTime As Long, Cache As String, JSON As String, Model As String, JS As Object
   Dim TimeStart As Double, TimeEnd As Double, FileName As String
   RepName="MicrosoftDI"
   'RepName="PDFTEXT"   'uncomment this line if you want Advanced Zone Locator to use Text
   While pXDoc.Representations.Count>0
      If pXDoc.Representations(0).Name=RepName Then Exit Sub 'We already have Microsoft OCR text, no need to call again.
      pXDoc.Representations.Remove(0) ' remove all OCR results from XDocument
   Wend
   EndPoint=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceEndpoint").Value 'The Microsoft Azure Cloud URL
   Key=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceKey").Value   'Key to use Microsoft Cognitive Services
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value
   JSON=Cache_Load(pXDoc,"MicrosoftDI_JSON")
   pXDoc.Representations.Create(RepName)
   If JSON="" Then
      If pXDoc.CDoc.SourceFiles.Count=1 Then 'Does the XDoc contain 1 or more image files.
         FileName=pXDoc.CDoc.SourceFiles(0).FileName
      Else
         FileName = XDocument_ConvertToMultipageTIFF(pXDoc,False) 'We can send only one image to Microsoft. If we have multiple images,we merge them into a multipage TIFF.
      End If
      StartTime=Timer
      JSON=MicrosoftDI_REST(FileName,Model,EndPoint,Key,10)
      TimeEnd=Timer
      If pXDoc.CDoc.SourceFiles.Count>1 Then Kill FileName 'delete temp multipage tiff
      If TimeEnd<TimeStart Then 'Store time in milliseconds that Microsoft took to read document
         pXDoc.TimeOCR = CLng(1000 * (86400 - TimeStart + TimeEnd)) ' 86400=24*60^2 = seconds/day. needed if the job started before midnight and finished after midnight
      Else
         pXDoc.TimeOCR = CLng(1000 * (TimeEnd - TimeStart))  ' this is in milliseconds (accuracy of 1/18th of a second)
      End If

      If pXDoc.XValues.ItemExists("MicrosoftDI_Time") Then pXDoc.XValues.Delete("MicrosoftDI_Time")
      pXDoc.XValues.Add("MicrosoftDI_Time",CStr(Timer-StartTime),True)
      Cache_Save(pXDoc,"MicrosoftDI_JSON",JSON)
   End If
   Set JS= JSON_Parse(JSON)
   MicrosoftDI_AddWords(pXDoc, JS, 0)
End Sub

Public Function MicrosoftDI_REST(ImageFileName As String, Model As String, EndPoint As String, Key As String,Retries As Long) As String
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

Public Sub MicrosoftDI_AddWords(pXDoc As CscXDocument, JS As Object, PageOffset As Long)
   Dim P As Long, W As Long, Key As String, Confidences As String, Word As CscXDocWord, Units As String, XRes As Double, YRes As Double
   For P=0 To JS("js.analyzeResult.pages._count")-1
      Units=JS("JS.analyzeResult.Pages(" & CStr(P) & ").unit")
      If Units="inch" Then
         XRes=pXDoc.CDoc.Pages(P).XRes
         YRes=pXDoc.CDoc.Pages(P).XRes
      End If
      For W=0 To JS("js.analyzeResult.pages(" & P & ").words._count")-1   'format
         Key="js.analyzeResult.pages(" & P & ").words(" & W & ")"
         Set Word = New CscXDocWord
         Word.Text=JSON_Unescape(JS(Key & ".content"))
         Word.PageIndex=P
         JSON_Polygon2Rectangle(JS,Key,Word,Units,XRes,YRes)
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

Public Sub MicrosoftDI_AddTables(pXDoc As CscXDocument, JS As Object, PageOffset As Long)

End Sub

Public Sub MicrosoftDI_AddTable(pXDoc As CscXDocument, JS As Object, Table As CscXDocTable, t As Long)
   Dim Row As CscXDocTableRow, R As Long, C As Long, CellIndex As Long, Cell As CscXDocTableCell, W As Long, Words As CscXDocWords, P As Long, Key As String, BR As Long, BRKey As String
   Dim rowCount As Long, columnCount As Long, Units As String, XRes As Double, YRes As Double
   Table.Rows.Clear
   rowCount =CLng(JS("js.analyzeResult.tables(" & t & ").rowCount"))
   While Table.Rows.Count<rowCount
      Table.Rows.Append
   Wend
   columnCount = CLng(JS("js.analyzeResult.tables(" & t & ").columnCount"))
   For CellIndex =0 To rowCount*columnCount-1
      Key="js.analyzeResult.tables(" & t & ").cells(" & CellIndex & ")"
      R=CLng(JS(Key & ".rowIndex"))
      C=CLng(JS(Key & ".columnIndex"))
      If C<Table.Columns.Count Then
         Set Cell=Table.Rows(R).Cells(C)
         'Cell.Text=JSON_Unescape(JS(Key & ".content"))
         For BR = 0 To CLng(JS(Key & ".boundingRegions._count"))-1
            BRKey = Key & ".boundingRegions(" & BR & ")"
            P =CLng(JS(BRKey & ".pageNumber"))-1
            Units=JS("JS.analyzeResult.Pages(" & CStr(P+1) & ").unit")
            If Units="inch" Then
               XRes=pXDoc.CDoc.Pages(P).XRes
               YRes=pXDoc.CDoc.Pages(P).XRes
            End If
            JSON_Polygon2Rectangle(JS,Key,Cell, Units, XRes, YRes)
            Set Words = pXDoc.GetWordsInRect(P,Cell.Left,Cell.Top, Cell.Width, Cell.Height)
            For W=0 To Words.Count-1
               Cell.AddWordData(Words(W))
            Next
         Next
      End If
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
   Dim P As Long, FileFormat As Long, ColorFormat As Long

   For P = 0 To pXDoc.CDoc.Pages.Count - 1
      'Derive new filename from existing name - just replace extension with .tif
      With pXDoc.CDoc.Pages(P)
         Set SourceImg=pXDoc.CDoc.Pages(.IndexOfSourceFile).GetImage()
         If P = 0 Then
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

Public Function Cache_Load(pXDoc As CscXDocument, RepName As String) As String
   Dim R As Long, Model As String, CacheFileName As String
   For R=0 To pXDoc.Representations.Count-1
      If pXDoc.Representations(R).Name=RepName Then Return pXDoc.Representations(R).Text
   Next
   Model=Project.ScriptVariables.ItemByName("MicrosoftDocumentIntelligenceModel").Value
   CacheFileName=Replace(pXDoc.FileName,".xdc", "." & Model & ".json")
   If Dir(CacheFileName)<>"" Then Return File_Load(CacheFileName)
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
   Rectangle.Left=  Min(CDouble(JS(Key & ".polygon(0)")),CDouble(JS(Key & ".polygon(6)")))
   Rectangle.Width= Max(CDouble(JS(Key & ".polygon(2)")),CDouble(JS(Key & ".polygon(4)")))-Rectangle.Left
   Rectangle.Top =  Min(CDouble(JS(Key & ".polygon(1)")),CDouble(JS(Key & ".polygon(3)")))
   Rectangle.Height=Max(CDouble(JS(Key & ".polygon(5)")),CDouble(JS(Key & ".polygon(7)")))-Rectangle.Top
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
   Set AF=Project.FieldFormatters.ItemByName("DefaultAmountFormatter")
   AF.FormatField(F)
   Return F.DoubleValue
End Function

