using Microsoft.SharePoint.Client;
using System;
using System.Data;
using System.IO;
using System.Net;
using System.Security;



namespace SharePointOnline
{
    public class SharepointIntegration
	{
		public static string sharePointUsername;

		public static string sharePointPassword { get; set; }
		static SharepointIntegration()
		{
			SharepointIntegration.sharePointUsername = string.Empty;
			SharepointIntegration.sharePointPassword = string.Empty;
		}

        public SharepointIntegration()
        {

        }

        public SharepointIntegration(string username, string password)
        {
			SharepointIntegration.sharePointUsername = username;
			SharepointIntegration.sharePointPassword = password;
		}

		public void SetCredentials(string username, string password)
		{
			SharepointIntegration.sharePointUsername = username;
			SharepointIntegration.sharePointPassword = password;
		}

		public string UploadFile(string sharepointRootPath, string sharepointFolderPath, string fileName, bool isOverRide, string filepath)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					Web web = clientContext.Site.RootWeb;
					FileCreationInformation newFile = new FileCreationInformation();
					newFile.ContentStream = (Stream) new MemoryStream(System.IO.File.ReadAllBytes(filepath));
					newFile.Url = string.Concat(sharepointRootPath, sharepointFolderPath, fileName);
					newFile.Overwrite = isOverRide;
					List docs = web.Lists.GetByTitle("Documents");
                    clientContext.Load(docs);
                    Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);
                    clientContext.Load(uploadFile);
                    clientContext.ExecuteQuery();
                    return "Success";
				}
					
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public string DownloadFile(string sharepointRootPath, string sharepointFolderPath, string sharepointFileName, string outputDirectory)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					clientContext.Load(list.RootFolder.Folders);
					Web web = clientContext.Web;
					Folder folders = web.GetFolderByServerRelativeUrl(sharepointFolderPath);
					clientContext.Load(folders);
					clientContext.Load(folders.Files);
					clientContext.ExecuteQuery();
					FileCollection fileCollection = folders.Files;
					foreach (Microsoft.SharePoint.Client.File file in fileCollection)
					{
						if (file.Name == sharepointFileName)
						{
							string sharepointPath = string.Concat(sharepointRootPath, sharepointFolderPath, "/", sharepointFileName);
							string localFilePath = string.Concat(outputDirectory, "/", sharepointFileName);
							WebClient webClient = new WebClient();
							webClient.Credentials = clientContext.Credentials;
							webClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
							webClient.Headers.Add("X-IDCRL_ACCEPTED", "t");
							webClient.DownloadFile(sharepointPath, localFilePath);
							return "Success";
						}
					}
					return "File not found.";
				}
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public string DownloadFilesFromFolder(string sharepointRootPath, string sharepointFolderPath, string outputDirectory)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					int filecount = 0;
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					clientContext.Load(list.RootFolder.Folders);
					Web web = clientContext.Web;
					Folder folders = web.GetFolderByServerRelativeUrl(sharepointFolderPath);
					clientContext.Load(folders);
					clientContext.Load(folders.Files);
					clientContext.ExecuteQuery();
					FileCollection fileCollection = folders.Files;
					foreach (Microsoft.SharePoint.Client.File file in fileCollection)
					{
						FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, file.ServerRelativeUrl);
						clientContext.ExecuteQuery();
						var filePath = outputDirectory + file.Name;
						using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
						{
							fileInfo.Stream.CopyTo(fileStream);
							filecount++;
						}
					}
					if (filecount != 0)
					{
						return "Success";
					}
					else
					{
						return "No files found.";
					}
				}
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public string CreateFolder(string sharepointRootPath, string sharepointFolderCreationPath, string folderName)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					Web web = clientContext.Web;
					Folder folder = web.GetFolderByServerRelativeUrl(sharepointFolderCreationPath);
					clientContext.Load(folder);
					string folderPath = string.Concat(sharepointRootPath, sharepointFolderCreationPath);
					folder.Folders.Add(folderName);
					folder.Update();
					clientContext.ExecuteQuery();
					return "Success";
				}
			}
			catch (Exception ex)
			{
				return "Error occured while creating the folder";
			}
		}

		public string DeleteFile(string sharepointRootPath, string sharepointFileDeletePath, string fileName)
		{
			try
			{
				using(ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					int count = 0;
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					Web web = clientContext.Site.RootWeb;
					Folder folder = web.GetFolderByServerRelativeUrl(sharepointFileDeletePath);
					clientContext.Load(folder);
					clientContext.Load(folder.Files);
					clientContext.ExecuteQuery();
					FileCollection fcollection = folder.Files;
					foreach (Microsoft.SharePoint.Client.File file in fcollection)
					{
						if (file.Name == fileName)
						{
							file.DeleteObject();
							clientContext.ExecuteQuery();
							count++;
							break;
						}
					}
					if(count == 0)
					{
						return "File not found";
					}
					else
					{
						return "Success";
					}
				}
			}
			catch(Exception ex)
			{
				return ex.Message.ToString();
			}			
		}

		public string DeleteFolder(string sharepointRootPath, string sharepointDeleteFolderPath, string folderName)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					bool folderFound = false;
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					clientContext.Load(list.RootFolder.Folders);
					Web web = clientContext.Site.RootWeb;
					Folder folders = web.GetFolderByServerRelativeUrl(sharepointDeleteFolderPath);
					clientContext.Load(folders);
					clientContext.Load(folders.Folders);
					clientContext.ExecuteQuery();
					FolderCollection folderCollection = folders.Folders;
					foreach (Microsoft.SharePoint.Client.Folder folder in folderCollection)
					{
						if (folder.Name == folderName)
						{
							folder.DeleteObject();
							clientContext.ExecuteQuery();
							folderFound = true;
							break;
						}
					}
					if (folderFound)
					{
						return "Success";
					}
					else
					{
						return "Folder not found.";
					}
				}
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public string GetFileInfo(string sharepointRootPath, string sharepointFolderPath, string fileName, string pathToSave, string localFileName)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					DataTable dtFileDetails = new DataTable();
					dtFileDetails.Columns.Add("File Name", typeof(string));
					dtFileDetails.Columns.Add("Checked out by user", typeof(string));
					dtFileDetails.Columns.Add("Modified by", typeof(string));
					dtFileDetails.Columns.Add("Last Modified Time", typeof(string));
					dtFileDetails.Columns.Add("Created time", typeof(string));
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					Web web = clientContext.Site.RootWeb;
					Folder folder = web.GetFolderByServerRelativeUrl(sharepointFolderPath);
					clientContext.Load(folder);
					clientContext.Load(folder.Files);
					clientContext.ExecuteQuery();
					FileCollection foldercollection = folder.Files;
					foreach (Microsoft.SharePoint.Client.File file in foldercollection)
					{
						if (file.Name == fileName)
						{
							clientContext.Load(file);
							clientContext.Load(file.ListItemAllFields);
							clientContext.ExecuteQuery();
							DataRow rowData = dtFileDetails.NewRow();
							rowData["File Name"] = file.Name.ToString();
							FieldUserValue checkedoutuser = (FieldUserValue)file.ListItemAllFields.FieldValues["CheckoutUser"];
							if(checkedoutuser != null)
							{
								
								rowData["Checked out by user"] = checkedoutuser.LookupValue;
							}
							rowData["Modified by"] = file.ListItemAllFields.FieldValues["Modified_x0020_By"].ToString() != null ? file.ListItemAllFields.FieldValues["Modified_x0020_By"].ToString() : "";
							rowData["Last Modified Time"] = file.TimeLastModified.ToString();
							rowData["Created time"] = file.TimeCreated.ToString();
							dtFileDetails.Rows.Add(rowData);
							clientContext.ExecuteQuery();
						}
					}
					string response = Convert(dtFileDetails, pathToSave, localFileName);
					return response;
				}
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public string GetFolderInfo(string sharepointRootPath, string sharepointFolderPath, string pathToSave, string localFileName)
		{
			try
			{
				using (ClientContext clientContext = new ClientContext(sharepointRootPath))
				{
					DataTable dtFolderDetails = new DataTable();
					dtFolderDetails.Columns.Add("Folder Name", typeof(string));
					SecureString passWord = new SecureString();
					foreach (char c in sharePointPassword.ToCharArray()) passWord.AppendChar(c);
					clientContext.Credentials = new SharePointOnlineCredentials(sharePointUsername, passWord);
					List list = clientContext.Web.Lists.GetByTitle("Documents");
					clientContext.Load(list);
					clientContext.Load(list.RootFolder);
					Web web = clientContext.Site.RootWeb;
					Folder folders = web.GetFolderByServerRelativeUrl(sharepointFolderPath);
					clientContext.Load(folders);
					clientContext.Load(folders.Folders);
					clientContext.ExecuteQuery();
					FolderCollection foldercollection = folders.Folders;
					foreach (Microsoft.SharePoint.Client.Folder folder in foldercollection)
					{
						DataRow rowData = dtFolderDetails.NewRow();
						rowData["Folder Name"] = folder.Name.ToString();
						dtFolderDetails.Rows.Add(rowData);
					}
					string response = Convert(dtFolderDetails, pathToSave, localFileName);
					return response;
				}
			}
			catch (Exception ex)
			{
				return ex.Message.ToString();
			}
		}

		public static string Convert(DataTable dtData, string pathToSave, string localFileName)
		{
			try
			{
				string csv = string.Empty;
				foreach (DataColumn column in dtData.Columns)
				{
					csv += column.ColumnName + ',';
				}
				csv += "\r\n";
				foreach (DataRow row in dtData.Rows)
				{
					foreach (var rowdata in row.ItemArray)
					{
						csv += rowdata.ToString().Replace(",", ";") + ',';
					}
					csv += "\r\n";
				}
				string folderPath = pathToSave + "\\" + localFileName + ".csv";
				System.IO.File.WriteAllText(folderPath, csv);
				return "Exported successfully";
			}
			catch(Exception ex)
			{
				return ex.Message.ToString();
			}
		}
	}
}

