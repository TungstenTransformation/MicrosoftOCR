Option Explicit

' Class script: NewClass1

Private Sub Document_AfterLocate(ByVal pxdoc As CASCADELib.CscXDocument, ByVal locatorname As String)
   Select Case locatorname
   Case "AZL"
      AZL_UseMicrosoftOCRConfidences(pxdoc,locatorname)
   End Select
End Sub

Sub AZL_UseMicrosoftOCRConfidences(ByVal pxdoc As CASCADELib.CscXDocument, ByVal locatorname As String)
   'This changes the AZL zonal confidences to the Microsoft OCR pagelevel word confidences.
   'The zone confidence is calculated using the word lengths as weights.
   Dim S As Long, SubFields As CscXDocSubFields, Confidences() As String, AF As ICscFieldFormatter, F As New CscXDocField
   Dim Words As CscXDocWords, W As Long, WordLength As Long
   If Not pxdoc.XValues.ItemExists("MicrosoftOCR_WordConfidences") Then Exit Sub ' there are no stored Microsoft OCR word confidences
   Set AF = Project.FieldFormatters.ItemByName("DefaultAmountFormatter") ' give the name of an Amount formatter with "." as decimal character to convert "0.992" to double.
   Confidences =Split(pxdoc.XValues.ItemByName("MicrosoftOCR_WordConfidences").Value,",")
   Set SubFields = pxdoc.Locators.ItemByName(locatorname).Alternatives(0).SubFields
   For S=0 To SubFields.Count-1
      With SubFields(S)
         .Confidence=0
         .LongTag=0 'use this customattribute to store the number of pagelevel OCR characters
         ' Find the OCR words that are inside the Zone
         Set Words=pxdoc.GetWordsInRect(.PageIndex,.Left,.Top,.Width,.Height)
         For W=0 To Words.Count-1
            F.Text=Confidences(Words(W).IndexOnDocument)
            AF.FormatField(F)
            WordLength=Len(Words(W).Text)
            .Confidence=.Confidence+ F.DoubleValue* WordLength   'Weight each word confidence by the word length
            .LongTag=.LongTag+WordLength
            .Text=Words.Text 'Ignore text from AZL zone profile, which get called if page-ocr text is not 100% within the zone.
         Next
         If .LongTag>0 Then .Confidence=.Confidence /.LongTag
      End With
   Next
End Sub
