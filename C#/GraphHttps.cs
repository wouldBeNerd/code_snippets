using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Linq;

//DOCS
//https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS
namespace SQL_to_GRAPH_v2_2021
{

    static class GraphHttps
    {

        //REPORTINGTEAM LINK
        private const string clientID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
        private const string aadInstance = "https://login.microsoftonline.com/{0}";
        private const string tenantID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"; //ends with onmicrosoft.com
        private const string resource = "https://graph.microsoft.com/";
        private const string clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
        private static ClientCredentialProvider authProvider;
        private static Microsoft.Graph.GraphServiceClient graphClient;

        public static void AUTH() {

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(clientID)
            .WithTenantId(tenantID)
            .WithClientSecret(clientSecret)
            .Build();

            authProvider = new ClientCredentialProvider(confidentialClientApplication);
            graphClient = new Microsoft.Graph.GraphServiceClient(authProvider);
            Console.WriteLine("AUTHED");

        }

        public static async Task<Drive> GetGroupDrive(string group_id, string drive_name)
        {
            var drives = await graphClient.Groups[group_id].Drives
                .Request()
                .GetAsync();


            Drive retDrive = new Drive();
            foreach (Drive drive in drives)
            {
                Console.WriteLine( $"id:{drive.Id} name:{drive.Name} " );
                if(drive.Name == drive_name){
                    retDrive = drive;
                    break;
                }
                
            }
            
            if(retDrive.Name != drive_name){
                throw new  InvalidOperationException("drive_name was not found");
            }

            return retDrive;

        }

        public static async Task<Group> GetGroupId(string group_name)
        {
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            var groups = await graphClient.Groups
                .Request()
                .Filter($"startswith(displayName, '{group_name}')")
                .Select("id, displayName")
                .GetAsync();

            
            Group retGroup = groups[0];
            foreach (Group group in groups)
            {   
                Console.WriteLine( $"id:{group.Id} name:{group.DisplayName} " );
                retGroup = group;
                break;
            }
            
            return retGroup;

        }

        public static async Task<Channel> GetChannel(string group_id, string channel_name)
        {
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            var channels = await graphClient.Teams[group_id].Channels
                .Request()
                .GetAsync();

            Channel retChannel = channels[0];
            foreach(Channel channel in channels){
                Console.WriteLine( $"id:{channel.Id} name:{channel.DisplayName} " );
                if(channel.DisplayName == channel_name){
                    retChannel = channel;
                    break;
                }
            }
            return retChannel;     
        }

        public static async Task<List<DriveItem>> GetDriveChildren(string drive_id)
        {
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            
            var children = await graphClient.Drives[drive_id].Root.Children
                         .Request()
                        .GetAsync();
            List<DriveItem> retChildren = children.ToList();
            return retChildren;

        }

        /// <summary>
        /// downloads a file from a teams >> group >> channel
        /// </summary>
        /// <param name="group_id"></param>
        /// <param name="channel_id"></param>
        /// <param name="file_name_teams">requires extensions to be included IE `ProductsReportData.xlsx`</param>
        /// <param name="write_to_path_file">where to write the file to IE `/downloads/ProductsReportData.xlsx`</param>
        /// <returns></returns>
        public static async Task DownloadFile(string group_id, string channel_id, string file_name_teams, string write_to_path_file){
            
            GraphServiceClient graphClient = new GraphServiceClient( authProvider );

            var driveItem = await graphClient.Teams[group_id].Channels[channel_id].FilesFolder
                .Request()
                .GetAsync();

            Console.WriteLine(driveItem);
            Console.WriteLine( $"id:{driveItem.Id} name:{driveItem.Name} fileCount:{driveItem.Folder.ChildCount} driveId:{driveItem.ParentReference.DriveId} " );

            var children = await graphClient.Drives[ driveItem.ParentReference.DriveId ].Items[driveItem.Id].Children
                .Request()
                .GetAsync();
            List<DriveItem> di_children = children.ToList();

            Console.WriteLine(di_children);


            string item_id = "";
            foreach (DriveItem d_item in di_children)
            {

                Console.WriteLine($"FILE name:{d_item.Name} id:{d_item.Id} ");
                if(d_item.Name == file_name_teams){
                    item_id = d_item.Id;
                }
                
            }
            if(item_id.Length > 0){
                Console.WriteLine($"item_id: {item_id}");

                var memory_stream = await graphClient.Drives[ driveItem.ParentReference.DriveId ].Items[$"{item_id}"].Content
                    .Request()
                    .GetAsync();

                    //You have to rewind the MemoryStream before copying
                    memory_stream.Seek(0, SeekOrigin.Begin);
                    using (FileStream fs = new FileStream(write_to_path_file, FileMode.OpenOrCreate))
                    {
                        memory_stream.CopyTo(fs);
                        fs.Flush();
                    }
                    Console.WriteLine("file written");

            }else{
                Console.WriteLine("item id not found");



            }
            // Console.WriteLine($"WAIT");

        }


        /// <summary>
        /// upload file to teams by group id channel id teams file name
        /// </summary>
        /// <param name="group_id"></param>
        /// <param name="channel_id"></param>
        /// <param name="file_name_teams">requires extensions to be included IE `ProductsReportData.xlsx`</param>
        /// <param name="local_path_file">destination of current file</param>
        /// <returns></returns>
        public static async Task UploadFile(string group_id, string channel_id, string file_name_teams, string local_path_file){
            
            GraphServiceClient graphClient = new GraphServiceClient( authProvider );

            byte[] file_bytes = await System.IO.File.ReadAllBytesAsync(local_path_file);
 

            using var stream = new System.IO.MemoryStream( file_bytes );

            var driveItem = await graphClient.Teams[group_id].Channels[channel_id].FilesFolder
                .Request()
                .GetAsync();

            Console.WriteLine(driveItem);
            Console.WriteLine( $"id:{driveItem.Id} name:{driveItem.Name} fileCount:{driveItem.Folder.ChildCount} driveId:{driveItem.ParentReference.DriveId} " );

            var children = await graphClient.Drives[ driveItem.ParentReference.DriveId ].Items[driveItem.Id].Children
                .Request()
                .GetAsync();
            List<DriveItem> di_children = children.ToList();

            Console.WriteLine(di_children);


            string item_id = "";
            foreach (DriveItem d_item in di_children)
            {

                Console.WriteLine($"FILE name:{d_item.Name} id:{d_item.Id} ");
                if(d_item.Name == file_name_teams){
                    item_id = d_item.Id;
                }
                
            }
            if(item_id.Length > 0){
                Console.WriteLine($"item_id: {item_id}");

                var response = await graphClient.Drives[ driveItem.ParentReference.DriveId ].Items[$"{item_id}"].Content
                    .Request()
                    .PutAsync<DriveItem>(stream);

                Console.WriteLine(response);
                Console.WriteLine("done fileupdate to teams");


            }
        }


        public static async Task CopyFileToTeams(string group_id, string item_id, string lcl_itempath)
        {
            var fileName = lcl_itempath;
            var currentFolder = System.IO.Directory.GetCurrentDirectory();
            var filePath = Path.Combine(currentFolder, fileName);

            FileStream fileStream = new FileStream(filePath, FileMode.Open);

            Microsoft.Graph.GraphServiceClient graphClient = new Microsoft.Graph.GraphServiceClient(authProvider);


            var response = await graphClient.Groups[group_id].Drive.Items[item_id].Content
                .Request()
                .PutAsync<DriveItem>(fileStream);

            Console.WriteLine(response);
            Console.WriteLine("downloaded file from teams");
        }

    }

}
