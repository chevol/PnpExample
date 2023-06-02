using Amazon.Lambda.Core;
using Amazon.S3.Model;
using Microsoft.SharePoint.Client;
using System.Security;
using UnbilledInventory.Interface;

namespace UnbilledInventory;

public class PNPSharepointUploader : ISharepointUploader
{
    public async Task UploadToSharepoint(string s3BucketName, string s3FileName, ILambdaLogger logger)
    {
        string sharepointSiteUrl = Environment.GetEnvironmentVariable("SharepointSiteUrl")!;
        string sharepointLibraryName = Environment.GetEnvironmentVariable("SharepointLibraryName")!;
        string sharepointClientSecret = Environment.GetEnvironmentVariable("SharepointClientSecret")!;
        string sharepointDefaultAppId = Environment.GetEnvironmentVariable("SharepointDefaultAppId")!;
        string sharepointRootFolder = Environment.GetEnvironmentVariable("SharepointRootFolder")!;
        string password = Environment.GetEnvironmentVariable("SharepointPassword")!;
        string sharepointFileName = s3FileName.Split('/')[2].Replace(".xlsx", $"_{Constants.DateTimeStamp}.xlsx").Replace("+"," ");
        string sharepointFolderName = s3FileName.Split('/')[0];

        try
        {
            SecureString passWord = new SecureString();
            foreach (var c in password)
            {
                passWord.AppendChar(c);
            }
            using (var cc = new PnP.Framework.AuthenticationManager().GetACSAppOnlyContext(sharepointSiteUrl, sharepointDefaultAppId, sharepointClientSecret))
            {
                cc.RequestTimeout = Timeout.Infinite;
                var s3Client = new Amazon.S3.AmazonS3Client();
                var getObjectRequest = new GetObjectRequest
                {
                    BucketName = s3BucketName,
                    Key = s3FileName
                };
                var getObjectResponse = await s3Client.GetObjectAsync(getObjectRequest);
                var folder = cc.Web.Lists.GetByTitle(sharepointLibraryName).RootFolder;
                var subFolder = folder;
                if (sharepointRootFolder != "")
                {
                    subFolder = folder.Folders.Add(sharepointRootFolder);
                    subFolder = subFolder.Folders.Add(sharepointFolderName);
                }
                else 
                {
                    subFolder = folder.Folders.Add(sharepointFolderName);
                }

                using (var stream = getObjectResponse.ResponseStream)
                {
                    await subFolder.UploadFileAsync(sharepointFileName, stream, true);
                }
            }
        }
        catch (Exception ex)
        {
            logger.LogError("Error uploading to Sharepoint");
            logger.LogError(ex.Message);
            logger.LogError(ex.StackTrace);
        }

    }
}
