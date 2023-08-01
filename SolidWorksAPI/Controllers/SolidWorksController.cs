using Azure.Storage.Blobs.Models;
using Microsoft.AspNetCore.Mvc;
using SolidWorks.Interop.sldworks;
using Azure.Storage.Blobs;
using IConfiguration = Microsoft.Extensions.Configuration.IConfiguration;

[Route("api/[controller]")]
[ApiController]
public class SolidWorksController : ControllerBase
{
    private readonly ISldWorks _swApp;
    private readonly Manager _manager;

    public SolidWorksController(ISldWorks swApp, Manager manager)
    {
        _swApp = swApp;
        _manager = manager;
    }

    [HttpGet("run")]
    public async Task<ActionResult> Run(string filePath)
    {
        _swApp.Visible = false;

        string downloadPath = Path.Combine(Directory.GetCurrentDirectory(), filePath.Replace('/', '\\'));
        await _manager.DownloadBlobFromAzure(filePath, downloadPath);

        await _manager.RotateAndTakeScreenshots(downloadPath);

        string azureOutPath = filePath.Replace("input.step", "out");
        await _manager.UploadFolder(Path.Combine(Path.GetDirectoryName(downloadPath)!, "out"), azureOutPath);

        return Ok();
    }
}

public class Manager
{
    private readonly ISldWorks _swApp;
    private readonly IConfiguration _configuration;

    public Manager(ISldWorks app, IConfiguration configuration)
    {
        _swApp = app;
        _configuration = configuration;
    }

    public async Task RotateAndTakeScreenshots(string stepFilePath)
    {
        int errors = 0;

        //Get import information
        ImportStepData swImportStepData = (ImportStepData)_swApp.GetImportFileData(stepFilePath);

        //If ImportStepData::MapConfigurationData is not set, then default to
        //the environment setting swImportStepConfigData; otherwise, override
        //swImportStepConfigData with ImportStepData::MapConfigurationData
        swImportStepData.MapConfigurationData = true;

        Console.Write(stepFilePath);
        //Import the STEP file
        PartDoc swPart = (PartDoc)_swApp.LoadFile4(stepFilePath, "r", swImportStepData, ref errors);
        Console.Write(errors);
        ModelDoc2 swModel = (ModelDoc2)swPart;
        ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;

        //Run diagnostics on the STEP file and repair any bad faces
        errors = swPart.ImportDiagnosis(true, false, true, 0);

        swModel.ViewDisplayShaded();

        // Zoom to fit
        swModel.ViewZoomtofit2();

        string outPath = Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(stepFilePath), "out")).FullName;
        string screenshotPath = Path.Combine(outPath, "screen.bmp");

        swModel.SaveBMP(screenshotPath, 0, 0);

        _swApp.ExitApp();
    }

    public async Task DownloadBlobFromAzure(string blobName, string outPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outPath)!);

        var connectionString = _configuration["BlobStorage:ConnectionString"];
        var blobContainerName = _configuration["BlobStorage:ContainerName"];

        BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
        BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(blobContainerName);

        BlobClient blobClient = containerClient.GetBlobClient(blobName);

        BlobDownloadInfo download = await blobClient.DownloadAsync();

        using (FileStream downloadFileStream = File.OpenWrite(outPath))
        {
            await download.Content.CopyToAsync(downloadFileStream);
            downloadFileStream.Close();
        }
    }

    public async Task UploadFolder(string localPath, string azurePath)
    {
        var connectionString = _configuration["BlobStorage:ConnectionString"];
        var blobContainerName = _configuration["BlobStorage:ContainerName"];

        BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
        BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(blobContainerName);

        foreach (var path in Directory.GetFiles(localPath))
        {
            string blobName = Path.GetFileName(path);

            BlobClient blobClient = containerClient.GetBlobClient(azurePath + "/" + blobName);

            using (FileStream uploadFileStream = File.OpenRead(path))
            {
                await blobClient.UploadAsync(uploadFileStream, true);
                uploadFileStream.Close();
            }

            Console.WriteLine($"Uploaded: {blobName}");
        }
    }
}
