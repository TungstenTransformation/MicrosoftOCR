# Microsoft Azure OCR for Kofax Transformation

Microsoft OCR is built on top of multiple deep learning models.
* JPEG, GIF, PNG, TIFF 50x50 to 4200x4200 pixels black&white or color.
* Automatic language detection.
* Handwritten and printed text.
* Cloud or on-premise via Docker Containers
* [ree tier: 20 calls/minute, 5k calls/month.
* S1 Tier 10: calls/second @ ~ 80$/1000 calls. See [Pricing]((https://azure.microsoft.com/en-gb/pricing/details/cognitive-services/computer-vision/)) for more.
* very simple integration into Kofax Transformation and Kofax Total Agility. No dlls, plugins or any other software required. Works also on KTA cloud.
## Configuration at Microsoft Azure
* Create a free account at [Microsoft Azure](https://azure.microsoft.com).
* Open [Azure Portal](https://portal.azure.com/#home).
* Click **Create a Resource**.
* search for **Computer Vision**.
* Click **Create**  
![Alt text](images/Computer%20Vision%20Create.png)
* Name your endpoint, select your Region, Pricing Tier and click **Review + Create**.
* Click on **Click here to manage keys**
* You will copy KEY1 or KEY2 and the Endpoint into Kofax Transformation.

## Configure Kofax Transformation

* Copy [script](Microsoft%20OCR.vb) into the Project Level Class.
* Rename the Default Page Recognition profile to **Microsoft OCR**.
* Add Two Script Variables to your project: 
    * **MicrosoftComputerVisionKey**
    * **MicrosoftComputerVisionEndpoint**
![Alt text](images/Script%20Variables.png)
* Classify a document
https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/concept-ocr




* [prerequisites](https://learn.microsoft.com/en-us/azure/cognitive-services/computer-vision/quickstarts-sdk/image-analysis-client-library?tabs=visual-studio%2C3-2&pivots=programming-language-rest-api#prerequisites)  

https://westus.dev.cognitive.microsoft.com/docs/services/computer-vision-v3-2/operations/56f91f2e778daf14a499f20d
