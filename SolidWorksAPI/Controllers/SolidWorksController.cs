using Azure.Storage.Blobs.Models;
using Microsoft.AspNetCore.Mvc;
using SolidWorks.Interop.sldworks;
using Azure.Storage.Blobs;
using IConfiguration = Microsoft.Extensions.Configuration.IConfiguration;
using System.Drawing;
using SolidWorks.Interop.swconst;
using System.Linq;


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
        _swApp.Visible = true;


        int errors = 0;

        string downloadPath = Path.Combine(Directory.GetCurrentDirectory(), filePath.Replace('/', '\\'));
        await _manager.DownloadBlobFromAzure(filePath, downloadPath);

        var dirName = Path.GetDirectoryName(downloadPath)!;
        Directory.CreateDirectory(dirName);

        string outPath = Path.Combine(dirName, "out");

        ModelDoc2 swModel = _manager.OpenStepFile(downloadPath);

        Console.WriteLine("1");

        ModelView view = (ModelView)swModel.ActiveView;

        Console.WriteLine("2");

        // view.FrameState = (int)swWindowState_e.swWindowMaximized;
        // _swApp.ActivateDoc2("Part14", false, errors);
        // Console.WriteLine(errors);

        swModel.ViewZoomtofit2();

        Console.WriteLine("3");


        // nt color = part.GetMaterialPropertyValues2((int)swInConfigurationOpts_e.swThisConfiguration, out _)[(int)swMaterialC];


        _manager.DraftAnalysis();
        Console.WriteLine("4");


        swModel.ForceRebuild3(true);

        Console.WriteLine("5");

        string[] screenshotPaths = _manager.RotateAndTakeScreenshots(swModel, outPath);

        Console.Write(string.Join(", ", screenshotPaths));

        Console.WriteLine("6");

        // _manager.ExitApp();

        string azureOutPath = filePath.Replace("input.step", "out");
        await _manager.UploadFolder(Path.Combine(outPath), azureOutPath);


        //return azure paths
        return Ok(
            screenshotPaths.Select(localPath => Path.Combine(azureOutPath, Path.GetFileName(localPath)).Replace("\\", "/")).ToArray()
        );
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

    public Boolean DraftAnalysis()
    {
        ModelDoc2 swModel = (ModelDoc2)_swApp.ActiveDoc;

        int[] colors = new int[] { 65280, 255, 65535, 16711680, 3446361 };

        swModel.MoldDraftAnalysis(0.05235987755983, 8, (object)colors, 127);

        return true;

        // bool status = swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);


        ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;


        /*         object[] positiveDraft = new object[] { 0, 1, 0 };
             object[] negativeDraft = new object[] { 1, 0, 0 };
             object[] noDraft = new object[] { 1, 1, 0 };
             object[] straddledFaces = new object[] { 0, 0, 1 }; */

        /*      int[] colors = new int[4];

             colors[0] = (255 << 16) | (0 << 8) | 0;    // Red color for positive draft
             colors[1] = (0 << 16) | (255 << 8) | 0;    // Green color for negative draft
             colors[2] = (0 << 16) | (0 << 8) | 255;    // Blue color for no draft
             colors[3] = (255 << 16) | (255 << 8) | 0;  // Yellow color for straddled faces

             int allValues = (int)(swDraftAnalysisShow_e.swDraftAnalysisShowNegative
                             | swDraftAnalysisShow_e.swDraftAnalysisShowNegativeSteep
                             | swDraftAnalysisShow_e.swDraftAnalysisShowPositive
                             | swDraftAnalysisShow_e.swDraftAnalysisShowPositiveSteep
                             | swDraftAnalysisShow_e.swDraftAnalysisShowStraddle
                             | swDraftAnalysisShow_e.swDraftAnalysisShowSurface); */

        // swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 1, null, 0);
    }

    public ModelDoc2 OpenStepFile(string stepFilePath)
    {
        int errors = 0;

        //Get import information
        ImportStepData swImportStepData = (ImportStepData)_swApp.GetImportFileData(stepFilePath);

        //If ImportStepData::MapConfigurationData is not set, then default to
        //the environment setting swImportStepConfigData; otherwise, override
        //swImportStepConfigData with ImportStepData::MapConfigurationData
        swImportStepData.MapConfigurationData = true;

        PartDoc swPart = (PartDoc)_swApp.LoadFile4(stepFilePath, "r", swImportStepData, ref errors);
        ModelDoc2 swModel = (ModelDoc2)swPart;

        //Run diagnostics on the STEP file and repair any bad faces
        // errors = swPart.ImportDiagnosis(true, false, true, 0);

        swModel.ViewDisplayShaded();

        // Zoom to fit
        swModel.ViewZoomtofit2();

        swModel.ViewConstraint();

        swModel.ViewDispCoordinateSystems();

        /*       swModel.ViewDisplayCurvature();
              swModel.ViewDisplayFaceted();
              swModel.ViewDispRefplanes();
       */
        return swModel;
    }

    public string[] RotateAndTakeScreenshots(ModelDoc2 swModel, string outDir)
    {

        ModelView view = (ModelView)swModel.ActiveView;

        /*         view.RotateAboutCenter(Math.PI / 2, Math.PI / 2);

                swModel.GraphicsRedraw2(); */

        Directory.CreateDirectory(outDir);


        var views = new (int viewID, string name)[]
        {
            (2, "backView"),
            (6, "bottomView"),
            (9, "dimetricView"),
            (1, "frontView"),
            (7, "isometricView"),
            (3, "leftView"),
            (4, "rightView"),
            (5, "topView"),
            (8, "trimetricView")
        };

        var paths = new string[views.Length];

        for (int i = 0; i < views.Length; i++)
        {
            var (viewID, name) = views[i];

            swModel.ShowNamedView2("", viewID);

            swModel.ViewZoomtofit2();

            string path = Path.Combine(outDir, name + ".bmp");
            swModel.SaveBMP(path, 0, 0);

            paths[i] = path;
        }

        return paths;
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

    public void ExitApp()
    {
        _swApp.ExitApp();
    }
}