# Microsoft Azure OCR for Kofax Transformation
##Downloads
* [Version 1.0.0](https://github.com/KofaxTransformation/MicrosoftOCR/releases/tag/1.0.0). Initial Release. Full page OCR in cloud or on site.

Microsoft OCR is built on top of multiple deep learning models.
* JPEG, GIF, PNG, TIFF 50x50 to 4200x4200 pixels, black&white or color.
* Automatic language detection. Supports [122+ languages](https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/language-support#optical-character-recognition-ocr).
* Handwritten and printed text.
* On Azure Cloud or [on-premise](https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/computer-vision-how-to-install-containers) via Docker Containers.
* Free tier: 20 calls/minute, 5k calls/month.
* S1 Tier 10: calls/second @ ~ 80$/1000 calls. See [Pricing]((https://azure.microsoft.com/en-gb/pricing/details/cognitive-services/computer-vision/)) for more.
* very simple integration into Kofax Transformation and Kofax Total Agility. No dlls, plugins or any other software required. Works also on KTA cloud.
## Configure Microsoft Azure
* Create a free account at [Microsoft Azure](https://azure.microsoft.com).
* Open [Azure Portal](https://portal.azure.com/#home).
* Click **Create a Resource**.
* search for **Computer Vision**.
* Click **Create**.
* Name your endpoint, select your Region, Pricing Tier and click **Review + Create**.
* Click on **Click here to manage keys**.
* You will need to copy either **KEY1** or **KEY2** and the **Endpoint** into Kofax Transformation.

## Configure Kofax Transformation

* Copy the [script](Microsoft%20OCR.vb) into the Project Level Class of your KT Project.
* Rename the Default Page Recognition profile to **Microsoft OCR**. *It doesn't matter what the OCR engine shown is, it will be ignored*.
* Add two Script Variables to your project in **Project/Configuration/Script Variables**: 
    * **MicrosoftComputerVisionKey**
    * **MicrosoftComputerVisionEndpoint**  
![Alt text](images/Script%20Variables.png)
* Paste in the Key and Endpoint that you copied from Microsoft Azure.
* Call Microsoft OCR by pressing F5 (Classify). (Pressing F4 will perform OCR without calling Microsoft.)

## How it works
In KTM and KTA runtime. Kofax Transformation performs OCR on demand, either when Text Classification is required or when a locator needs text.
This script runs in the event **Document_BeforeClassify**, which occurs before KT ever tries to OCR the document. The script checks if you named a profile "Microsoft OCR". If so, it sends each page of the document to Microsoft and copies the words and coordinates into the XDocument. The XDocument now has an OCR layer called "Microsoft OCR", which will be used by the classifiers and locators - OCR won't be called again with another document.
In Project Builder or Design Studio, pressing F4 performs OCR with the built-in engines. To force it to use Microsoft OCR, press F5 (Classify) to send the document to Microsoft.

## Limitations and Potential Improvements
* force it to use a particular language. By default it supports multiple languages per document.
* not tested on PDF documents.
* not tested on multipage TIFF, but should work.
* will generate an error when you reach your license limit.  
* ignores word-level confidences.
* Ignores regions, which could be copied into KT paragraphs.

Open an [issue](https://github.com/KofaxTransformation/MicrosoftOCR/issues) if you find a bug or need a feature implemented.

## useful links
* https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/quickstarts-sdk/image-analysis-client-library?tabs=visual-studio%2C3-2&pivots=programming-language-rest-api#prerequisites
* https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/concept-ocr
* https://westus.dev.cognitive.microsoft.com/docs/services/computer-vision-v3-2/operations/56f91f2e778daf14a499f20d
