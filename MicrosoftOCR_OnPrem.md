# Microsoft OCR and DI on premise.
1. Install Docker Desktop on Windows along with Windows Subsystems for Linux.
1. Go to [portal.azure.com](https://portal.azure.com).
1. Click **"**Create a Resource** and select **Computer Vision**.
1. Click **create** and selection your region, name and pricing.
1. When it is created copy the **Endpoint**.
1. Click on **Manage Keys** and copy the Key 1. *The second key is for rotating keys*.
1. At the Windows Command Line type the following to download Computer Vision from Docker hub.
```cmd
docker pull mcr.microsoft.com/azure-cognitive-services/vision/read:latest
```
1. When downloading has finished create and run a Docker container using your **Endpoint** and **Key**.
```cmd
set Endpoint=************************
set Key=************************
docker run --rm -it -p 5001:5000 --memory 16g --cpus 8 mcr.microsoft.com/azure-cognitive-services/vision/read:latest mcr.microsoft.com/azure-cognitive-services/vision/read:latest Eula=accept Billing=%Endpoint% ApiKey=%Key%
```
1. You should now be able to call it using the REST API.
  * Type=POST 
  * URL=http://localhost:5001/vision/v3.2/read/analyze
  * Request Body = {"url":"https://bloximages.chicago2.vip.townnews.com/shorelinemedia.net/content/tncms/assets/v3/editorial/5/dd/5dd3a66b-2da3-5d7e-af6d-7458c3d59ebc/5db87a272ab76.image.jpg"}
  * Header **Content-Type:application/json**   is optional.
2. The Web Service call should return 202 (accepted) and the header **Operation-Location** contains the URL needed to retrieve the results.
3. Call the URL using GET in a loop until it returns 200 (OK).
4. The response body will contain the JSON of the OCR results.