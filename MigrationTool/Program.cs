using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Xml;
using log4net;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Specialized;

namespace MigrationTool
{
    class Program
    {
        // create an instance from ILog
        private static readonly log4net.ILog _log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static string[] migratedItemIds = null;
        static string[] folderIds = null;
        static void Main(string[] args)
        {

            try
            {
                _log.Info("Application Started");
                string siteSourceUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SourceSite"]);
                string targetSiteUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["TargetSite1"]);
                bool firstMigration = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["FirstMigration"]);

                if (firstMigration)
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate ()
                    {
                        using (SPSite sourceSite = new SPSite(siteSourceUrl))
                        {
                            // Get the Web From the Root Site
                            SPWeb sourceWeb = sourceSite.RootWeb;
                            StartMigration(sourceWeb);
                            folderIds = GetFolderIdsFromText();
                            MigrateFolders(sourceWeb);
                            _log.InfoFormat("Succeeded Migration to all SiteCollections");
                        }
                    });
                }
                else
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate ()
                    {
                        using (SPSite sourceSite = new SPSite(siteSourceUrl))
                        {
                            // Get the Web From the Root Site
                            SPWeb sourceWeb = sourceSite.RootWeb;
                            //StartMigration(sourceWeb);
                            folderIds = GetFolderIdsFromText();
                            MigrateFolders(sourceWeb);
                            _log.InfoFormat("Succeeded Migration to all SiteCollections");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception :  {0}", ex.Message);
            }
            Console.ReadLine();

        }

        /// <summary>
        /// Copy Role Assignments from source site to target site
        /// </summary>
        /// <param name="sourceWeb"></param>
        /// <param name="destinationWeb"></param>
        public static void CopyWebRoleAssignments(SPWeb sourceWeb, SPWeb destinationWeb)
        {
            _log.WarnFormat("Starting Migrate Role Assignments from {0} to {1}", sourceWeb.Title, destinationWeb.Title);
            //Copy Role Assignments from source to destination web.
            foreach (SPRoleAssignment sourceRoleAsg in sourceWeb.RoleAssignments)
            {
                SPRoleAssignment destinationRoleAsg = null;
                SPGroup destinationGroup = null;
                SPGroup sourceGroup = null;
                //Get the source member object
                SPPrincipal member = sourceRoleAsg.Member;

                //Check if the member is a user 
                try
                {
                    SPUser sourceUser = (SPUser)member;
                    destinationWeb.EnsureUser(sourceUser.LoginName);
                    SPUser destinationUser = destinationWeb.AllUsers[sourceUser.LoginName];
                    _log.InfoFormat("Member is User");
                    if (destinationUser != null)
                    {
                        destinationRoleAsg = new SPRoleAssignment(destinationUser);
                    }
                }
                catch (Exception ex)
                {
                    if (destinationRoleAsg == null)
                    {
                        //Check if the member is a group
                        try
                        {
                            sourceGroup = (SPGroup)member;
                            // try to get group from destination Group
                            destinationGroup = destinationWeb.SiteGroups[sourceGroup.Name];
                        }
                        catch (Exception ex1)
                        {
                            // exception happen we can not get the group
                            sourceGroup = (SPGroup)member;
                            destinationWeb.SiteGroups.Add(sourceGroup.Name, sourceGroup.Owner, null, sourceGroup.Description);
                            _log.InfoFormat("Add Group: {0}  To site: {1}", sourceGroup.Name, destinationWeb.Title);
                        }
                    }
                    destinationGroup = destinationWeb.SiteGroups[sourceGroup.Name];
                    foreach (SPUser user in sourceGroup.Users)
                    {
                        destinationGroup.AddUser(user);
                        _log.InfoFormat("Add User: {0}  To Group: {1}", user.Name, sourceGroup.Name);
                    }
                    destinationRoleAsg = new SPRoleAssignment(destinationGroup);
                }

                if (destinationRoleAsg != null)
                {
                    foreach (SPRoleDefinition sourceRoleDefinition in sourceRoleAsg.RoleDefinitionBindings)
                    {
                        try
                        {
                            if (!destinationRoleAsg.RoleDefinitionBindings.Contains(destinationWeb.RoleDefinitions[sourceRoleDefinition.Name]))
                            {
                                if (sourceRoleDefinition.Name != "Limited Access")
                                {
                                    destinationRoleAsg.RoleDefinitionBindings.Add(destinationWeb.RoleDefinitions[sourceRoleDefinition.Name]);
                                    _log.InfoFormat("Add RoleDefinitionBindings: {0}  To Group: {1}", sourceRoleDefinition.Name, destinationRoleAsg.Member.Name);
                                }
                            }
                        }
                        catch (Exception exDefinition)
                        {

                            destinationWeb.RoleDefinitions.Add(sourceRoleDefinition);
                            _log.InfoFormat("Add PermissionLevel: {0}  To Site: {1}", sourceRoleDefinition.Name, destinationWeb.Title);

                            if (!destinationRoleAsg.RoleDefinitionBindings.Contains(destinationWeb.RoleDefinitions[sourceRoleDefinition.Name]))
                            {
                                if (sourceRoleDefinition.Name != "Limited Access")
                                {
                                    destinationRoleAsg.RoleDefinitionBindings.Add(destinationWeb.RoleDefinitions[sourceRoleDefinition.Name]);
                                    _log.InfoFormat("After Adding Permission level Add RoleDefinitionBindings: {0}  To Group: {1}", sourceRoleDefinition.Name, destinationRoleAsg.Member.Name);
                                }
                            }
                        }
                    }
                    if (destinationRoleAsg.RoleDefinitionBindings.Count > 0)
                    {
                        try
                        {
                            destinationWeb.RoleAssignments.Add(destinationRoleAsg);
                            _log.InfoFormat("Succeeded to Add Role Assignment from Site: {0}  to Site: {1}", sourceWeb.Title, destinationWeb.Title);
                        }
                        catch (ArgumentException)
                        {
                            _log.FatalFormat("Failed To add Role Assignment  from Site: {0}  to Site: {1} - Group Or User: {2} ", sourceWeb.Title, destinationWeb.Title, destinationRoleAsg.Member.Name);
                        }
                    }
                }
            }
            destinationWeb.Update();
            _log.WarnFormat("Migrate Role Assignments from Site: {0} To Site: {1} has been finished", sourceWeb.Title, destinationWeb.Title);
        }



        /// <summary>
        /// Copy Site Columns from source site to target site
        /// </summary>
        /// <param name="sourceWeb"></param>
        /// <param name="destinationWeb"></param>
        public static void CopySiteColumns(SPWeb sourceWeb, SPWeb destinationWeb)
        {
            _log.WarnFormat("Starting Migrate Site Column from Site: {0} To Site: {1}", sourceWeb.Title, destinationWeb.Title);
            SPFieldCollection siteColumns = sourceWeb.Fields;
            foreach (SPField field in siteColumns)
            {
                if (!destinationWeb.Fields.Contains(field.Id))
                {
                    try
                    {
                        string fieldXml = field.SchemaXml;
                        if (fieldXml.Contains("Version"))
                        {
                            XmlDocument docXml = new XmlDocument();
                            docXml.LoadXml(fieldXml);
                            XmlElement root = docXml.DocumentElement;
                            root.RemoveAttribute("Version");
                            fieldXml = docXml.OuterXml;
                        }
                        destinationWeb.Fields.AddFieldAsXml(fieldXml);
                        _log.InfoFormat("Succeeded to add siteColumn: {0} to  Site: {1}", field.StaticName, destinationWeb.Title);
                    }
                    catch (Exception ex)
                    {
                        WriteFailedContenttypes(ex.Message);
                        _log.FatalFormat("Failed to add siteColumn: {0} to  Site: {1} - Exception Message: {2} ", field.StaticName, destinationWeb.Title, ex.Message);
                    }

                }
            }

            _log.WarnFormat("Migrate Site Columns from Site: {0} To Site: {1} has been finished", sourceWeb.Title, destinationWeb.Title);
            // After Mapping SiteColumns Now we Can Map Content Types
            CopyContentTypes(sourceWeb, destinationWeb);
        }

        /// <summary>
        /// Copy Content Types from source site to target site
        /// </summary>
        /// <param name="sourceWeb"></param>
        /// <param name="destinationWeb"></param>
        public static void CopyContentTypes(SPWeb sourceWeb, SPWeb destinationWeb)
        {
            _log.WarnFormat("Starting Migrate Content types from Site: {0} to Site: {1}", sourceWeb.Title, destinationWeb.Title);
            SPContentTypeCollection contentTypeCollection = sourceWeb.ContentTypes;
            foreach (SPContentType contentType in contentTypeCollection)
            {
                try
                {
                    if (destinationWeb.ContentTypes[contentType.Name] == null)
                    {
                        destinationWeb.ContentTypes.Add(contentType);
                        _log.InfoFormat("Succeeded to Add Content Type: {0} to Site: {1} ", contentType.Name, destinationWeb.Title);
                        destinationWeb.Update();
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        if (destinationWeb.ContentTypes[contentType.Name] == null)
                        {
                            SPContentType contentTp = new SPContentType(contentType.Parent, destinationWeb.ContentTypes, contentType.Name);
                            foreach (SPField SF in contentType.Fields)
                            {
                                SPFieldLink field = new SPFieldLink(SF);

                                contentTp.FieldLinks.Add(field);
                            }

                            destinationWeb.ContentTypes.Add(contentTp);
                            destinationWeb.Update();
                            _log.InfoFormat("Succeeded to Add Content Type: {0} to Site: {1} ", contentType.Name, destinationWeb.Title);

                        }
                    }
                    catch (Exception exception)
                    {
                        _log.FatalFormat("Failed to Add Content Type: {0} to Site: {1}, Exception: {2}", contentType.Name, destinationWeb.Title, exception.Message);
                        WriteFailedContenttypes(string.Format("Failed to Add Content Type: {0} to Site: {1}, Exception: {2}", contentType.Name, destinationWeb.Title, exception.Message));
                    }

                }
            }
            _log.WarnFormat("Migrate Content types from Site: {0} to Site: {1} has been finished", sourceWeb.Title, destinationWeb.Title);
        }

        /// <summary>
        /// Migrate Folders And their content from source site to target site
        /// </summary>
        /// <param name="sourceWeb"></param>
        public static void MigrateFolders(SPWeb sourceWeb)
        {
            try
            {
                _log.InfoFormat("We in MigrateFolders");
                string docLibraryName = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["DocLibName"]);
                SPQuery query = new SPQuery();
                query.Query = "<OrderBy><FieldRef Name ='ID' Ascending = 'True'/></OrderBy>";
                //query.ViewAttributes = "Scope='RecursiveAll'";
                SPListItemCollection itemCollection = sourceWeb.Lists[docLibraryName].GetItems(query);
                int number = 0;
                int siteCollectionNum = 1;
                int numberOfFoldersPerSiteCollection = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["numberOfFoldersPerSiteCollection"]);
                int numberOfTargetSiteCollections = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["numberOfTargetSiteCollections"]);
                //get migrated Ids from text 
                migratedItemIds = ReadMigratedIds();

                foreach (SPListItem item in itemCollection)
                {
                    string targetSiteUrl = string.Empty;

                    string[] foldervalues = GetValuesForMigratedLine(folderIds.Where(c => c.Contains(Convert.ToString(item.ID))).ToList(), item.ID);
                    bool folderInSpecificSite = foldervalues[0] == Convert.ToString(item.ID);
                    if (!folderInSpecificSite)
                    {
                        if (numberOfTargetSiteCollections != 0)
                        {
                            string targetKey = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["firstPartTarget"]) + siteCollectionNum;
                            targetSiteUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings[targetKey]);
                        }

                    }
                    else
                    {
                        targetSiteUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SpecificTargetSite"]);
                    }
                    try
                    {
                        if (!string.IsNullOrEmpty(targetSiteUrl))
                        {
                            using (SPSite site = new SPSite(targetSiteUrl))
                            {
                                SPWeb targetWeb = site.RootWeb;

                                int level = findLevel(item.Url);
                                if (item.FileSystemObjectType.ToString() == "Folder")
                                {
                                    string folderUrl = string.Empty;
                                    string[] folderurlSplitted = item.Folder.ServerRelativeUrl.Split('/');
                                    for (int i = 1; i < folderurlSplitted.Length - 1; i++)
                                    {
                                        folderUrl += "/" + folderurlSplitted[i];
                                    }
                                    if (targetWeb.ServerRelativeUrl != "/")
                                        folderUrl = targetWeb.ServerRelativeUrl + folderUrl;

                                    bool migratedId = false;
                                    if (migratedItemIds != null)
                                    {
                                        string[] values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(item.ID))).ToList(), item.ID);
                                        migratedId = values[0] == Convert.ToString(item.ID);
                                    }

                                    if (!migratedId)
                                    {
                                        try
                                        {
                                            _log.WarnFormat("Migration Starting for ItemUrl: {0} from Site: {1} to Site {2}", item.Url, sourceWeb.Title, targetWeb.Title);
                                            SPListItem itemCreation = targetWeb.Lists[docLibraryName].Items.Add(folderUrl, item.FileSystemObjectType, Convert.ToString(item["FileLeafRef"]));
                                            itemCreation.Update();
                                            _log.InfoFormat("Create folder: {0} and the Url is {1}", Convert.ToString(item["FileLeafRef"]), folderUrl);

                                            foreach (SPField field in item.Fields)
                                            {
                                                if (field.Group != "_Hidden" && !field.Hidden && !field.Sealed && field.InternalName != "ID" && !field.ReadOnlyField)
                                                {
                                                    itemCreation[field.InternalName] = item[field.InternalName];
                                                }
                                            }
                                            itemCreation.Update();
                                            _log.InfoFormat("Succeeded to add MetaData to Folder: {0} in Site: {1}", Convert.ToString(item["FileLeafRef"]), targetWeb.Title);
                                            // write Succeeded Id
                                            WriteMigratedIds(item.ID, itemCreation.Folder.Url);
                                            if (item.Folder.Files.Count > 0)
                                                GenerateFiles(item.Folder, itemCreation.Folder, targetWeb, sourceWeb);
                                            if (item.Folder.SubFolders.Count > 0)
                                            {
                                                GenerateFolders(item.Folder.SubFolders, targetWeb, sourceWeb, docLibraryName);
                                            }
                                        }
                                        catch (Exception exMigrate)
                                        {
                                            WriteFailedMigratedIds(item.ID, item.Url);
                                            _log.FatalFormat("Exception Happened when Migrate Folder/File:{0} From Site: {1} to Site: {2} Exception: {3}", Convert.ToString(item["FileLeafRef"]), sourceWeb.Title, targetSiteUrl, exMigrate.Message);
                                        }
                                    }
                                    else
                                    {
                                        string[] values = null;


                                        values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(item.ID))).ToList(), item.ID);
                                        // migratedId = values[0] == Convert.ToString(item.ID);


                                        SPFolder folder = targetWeb.GetFolder(values[1]);
                                        if (item.Folder.Files.Count > 0)
                                            GenerateFiles(item.Folder, folder, targetWeb, sourceWeb);
                                        if (item.Folder.SubFolders.Count > 0)
                                        {
                                            GenerateFolders(item.Folder.SubFolders, targetWeb, sourceWeb, docLibraryName);
                                        }
                                    }

                                }
                                else if (item.FileSystemObjectType.ToString() == "File")
                                {
                                    bool migratedId = false;
                                    if (migratedItemIds != null)
                                    {
                                        string[] values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(item.ID))).ToList(), item.ID);
                                        migratedId = values[0] == Convert.ToString(item.ID);
                                    }
                                    if (!migratedId)
                                    {
                                        try
                                        {
                                            SPList list = targetWeb.Lists[docLibraryName];
                                            SPFolder folder = list.RootFolder;
                                            SPUser createdUser = targetWeb.EnsureUser(Convert.ToString(item["Author"]).Substring(Convert.ToString(item["Author"]).IndexOf("#") + 1));
                                            SPUser modifiedUser = targetWeb.EnsureUser(Convert.ToString(item["Editor"]).Substring(Convert.ToString(item["Editor"]).IndexOf("#") + 1));
                                            DateTime createdTime = item.File.TimeCreated;
                                            DateTime modifiedTime = item.File.TimeLastModified;
                                            SPFile newFile = folder.Files.Add(item.Url, item.File.OpenBinaryStream(), createdUser, modifiedUser, createdTime, modifiedTime);
                                            _log.InfoFormat("Succeeded to add file : {0} and the  url: {1} In Site : {2}", newFile.Title, item.Url, sourceWeb.Title);
                                            foreach (SPField field in item.Fields)
                                            {
                                                if (field.Group != "_Hidden" && !field.Hidden && !field.Sealed && field.InternalName != "ID" && !field.ReadOnlyField)
                                                    newFile.Item[field.InternalName] = item[field.InternalName];
                                            }
                                            newFile.Item.Update();
                                            _log.InfoFormat("Succeeded to add MetaData to File: {0} and the url: {1} in Site: {2}", newFile.Title, newFile.Url, targetWeb.Title);
                                            WriteMigratedIds(item.ID, string.Empty);
                                        }
                                        catch (Exception fileEx)
                                        {
                                            _log.FatalFormat("Exception Happened when Migrate Folder/File From Site: {0} to Site: {1} Exception: {2}", sourceWeb.Title, targetSiteUrl, fileEx.Message);
                                            WriteFailedMigratedIds(item.ID, item.Url);
                                        }
                                    }
                                }
                                if (level == 1 && !folderIds.Contains(Convert.ToString(item.ID)))
                                    number++;
                                if (number % numberOfFoldersPerSiteCollection == 0 && number != 0 && !folderIds.Contains(Convert.ToString(item.ID)))
                                    siteCollectionNum++;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        _log.FatalFormat("Exception Happened when Migrate Folder/File From Site: {0} to Site: {1} Exception: {2}", sourceWeb.Title, targetSiteUrl, ex.Message);
                    }
                }
            }

            catch (Exception ex)
            {
                _log.FatalFormat("Exception : {0}", ex.Message);
            }
        }

        /// <summary>
        /// Create Document in target web and add content type and fields to it
        /// </summary>
        /// <param name="sourceWeb"></param>
        /// <param name="targetWeb"></param>
        public static void CreateDocLibrary(SPWeb sourceWeb, SPWeb targetWeb)
        {
            string docLibraryName = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["DocLibName"]);
            try
            {
                _log.WarnFormat("Start Craeting Document Library: {0} in Site {1}", docLibraryName, targetWeb.Title);
                SPList sourceDoc = sourceWeb.Lists[docLibraryName];
                string description = sourceDoc.Description;

                targetWeb.AllowUnsafeUpdates = true;
                // Guid listGuid = null;
                try
                {
                    SPList listdestination = targetWeb.Lists[docLibraryName];
                }
                catch (Exception ex)
                {
                    WriteFailedContenttypes(ex.Message + "but it will be created");
                    try
                    {
                        Guid listGuid = targetWeb.Lists.Add(docLibraryName, description, SPListTemplateType.DocumentLibrary);
                        targetWeb.Update();
                        SPList newDoc = targetWeb.Lists[listGuid];
                        newDoc.ContentTypesEnabled = true;
                        _log.InfoFormat("Document Library: {0} Created In Site: {1} ", docLibraryName, targetWeb.Title);
                        foreach (SPContentType contentType in sourceDoc.ContentTypes)
                        {
                            if (newDoc.ContentTypes[contentType.Name] == null)
                            {
                                try
                                {
                                    newDoc.ContentTypes.Add(targetWeb.ContentTypes[contentType.Name]);
                                    _log.InfoFormat("Add Content type: {0} to Document Library: {1} ", contentType.Name, docLibraryName);
                                }
                                catch (Exception exCT)
                                {
                                    WriteFailedContenttypes("First Creation For Content Type " + exCT.Message);
                                    try
                                    {
                                        SPContentType contentTypeDelete = newDoc.ContentTypes[contentType.Name];
                                        newDoc.ContentTypes.Delete(contentTypeDelete.Id);
                                        newDoc.ContentTypes.Add(contentType);
                                        _log.InfoFormat("Add Content type: {0} to Document Library: {1} ", contentType.Name, docLibraryName);
                                    }
                                    catch (Exception exContenttype)
                                    {
                                        WriteFailedContenttypes("Second Creation For Content Type " + exContenttype.Message);
                                    }
                                }
                            }
                            else
                            {
                                try
                                {
                                    SPContentType contentTypeDelete = newDoc.ContentTypes[contentType.Name];
                                    newDoc.ContentTypes.Delete(contentTypeDelete.Id);
                                    newDoc.ContentTypes.Add(contentType);
                                    _log.InfoFormat("Add Content type: {0} to Document Library: {1} ", contentType.Name, docLibraryName);
                                }
                                catch (Exception exContenttype)
                                {
                                    WriteFailedContenttypes("Second Creation For Content Type " + exContenttype.Message);
                                }

                            }
                        }
                        _log.WarnFormat("Succeeded to Create Document Library : {0} and Add Content types to library In Site: {1}", docLibraryName, targetWeb.Title);

                    }
                    catch (Exception exCreateDoc)
                    {
                        _log.FatalFormat("Exception: {0}", exCreateDoc.Message);
                    }
                    // _log.FatalFormat("Exception Happened in Site: {0} when create Document Library: {1} and assign the Content Types, Exception Message : {2}", targetWeb.Title, docLibraryName, ex.Message);
                }

                try
                {
                    _log.InfoFormat("Check For Fields");
                    SPFieldCollection spFieldCollection = sourceDoc.Fields;
                    SPList newDoc = targetWeb.Lists[docLibraryName];
                    SPFieldCollection spFieldCollectionTarget = newDoc.Fields;
                    foreach (SPField field in spFieldCollection)
                    {
                        if (!spFieldCollectionTarget.ContainsField(field.InternalName))
                        {
                            try
                            {
                                newDoc.Fields.Add(field);
                                _log.InfoFormat("Add Field: {0} to Document Library In Site: {1}", field.InternalName, targetWeb.Name);
                            }
                            catch (Exception fieldEx)
                            {
                                _log.InfoFormat("Exception: {0}", fieldEx.Message);

                            }
                        }
                    }
                }
                catch (Exception fieldEx)
                {
                    _log.InfoFormat("Exception: {0}", fieldEx.Message);

                }

            }
            catch (Exception ex)
            {
                WriteFailedContenttypes(ex.Message);
                _log.FatalFormat("Exception Happened in Site: {0} when create Document Library: {1} and assign the Content Types, Exception Message : {2}", targetWeb.Title, docLibraryName, ex.Message);
                // targetWeb.Lists[docLibraryName].Delete();
                _log.InfoFormat("Delete the Document Library: {0}  and try create it again.", docLibraryName);
                // CreateDocLibrary(sourceWeb, targetWeb);
            }
        }

        /// <summary>
        /// find level for the item
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static int findLevel(string str)
        {
            int intCount = 0;
            while (str.IndexOf('/') != -1)
            {
                intCount++;
                str = str.Substring(str.IndexOf('/') + 1);
            }
            return intCount;
        }

        /// <summary>
        /// Start migration with calling three Methods CopyWebRoleAssignments - CopySiteColumns - CreateDocLibrary
        /// </summary>
        /// <param name="sourceWeb"></param>
        public static void StartMigration(SPWeb sourceWeb)
        {
            try
            {
                int numberOfSiteCollections = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["numberOfTargetSiteCollections"]);
                bool IsAnotherDoc = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["IsAnotherDoc"]);
                for (int i = 1; i <= numberOfSiteCollections + 1; i++)
                {
                    string targetKey = string.Empty;
                    string targetSiteUrl = string.Empty;
                    if (numberOfSiteCollections + 1 == i)
                    {
                        targetSiteUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SpecificTargetSite"]);
                    }
                    else
                    {
                        targetKey = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["firstPartTarget"]) + i;
                        targetSiteUrl = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings[targetKey]);
                    }


                    SPSecurity.RunWithElevatedPrivileges(delegate ()
                    {
                        using (SPSite site = new SPSite(targetSiteUrl))
                        {
                            SPWeb targetWeb = site.RootWeb;
                            targetWeb.AllowUnsafeUpdates = true;
                            _log.WarnFormat("Start Migration From Site {0} to Site {1}", sourceWeb.Title, targetWeb.Title);
                            if (!IsAnotherDoc)
                            {
                                CopyWebRoleAssignments(sourceWeb, targetWeb);
                                CopySiteColumns(sourceWeb, targetWeb);
                                CreateDocLibrary(sourceWeb, targetWeb);
                            }
                            else
                            {
                                CreateDocLibrary(sourceWeb, targetWeb);
                            }

                        }
                    });
                }
            }
            catch (Exception ex)
            {
                WriteFailedContenttypes(ex.Message);
                _log.FatalFormat("Exception : {0}", ex.Message);
            }
        }

        /// <summary>
        /// Add Files to Specific folder
        /// </summary>
        /// <param name="folderSource"></param>
        /// <param name="folder"></param>
        /// <param name="targetWeb"></param>
        /// <param name="sourceWeb"></param>
        public static void GenerateFiles(SPFolder folderSource, SPFolder folder, SPWeb targetWeb, SPWeb sourceWeb)
        {
            try
            {
                targetWeb.AllowUnsafeUpdates = true;
                SPFileCollection fileCollection = folderSource.Files;
                foreach (SPFile file in fileCollection)
                {
                    try
                    {
                        bool migratedId = false;
                        if (migratedItemIds != null)
                        {
                            string[] values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(file.Item.ID))).ToList(), file.Item.ID);
                            migratedId = values[0] == Convert.ToString(file.Item.ID);
                        }
                        if (!migratedId)
                        {
                            SPUser createdUser = targetWeb.EnsureUser(file.Author.LoginName);
                            SPUser modifiedUser = targetWeb.EnsureUser(file.ModifiedBy.LoginName);
                            DateTime createdTime = file.TimeCreated;
                            DateTime modifiedTime = file.TimeLastModified;
                            SPFileVersionCollection fileVersionCollection = file.Versions;
                            System.Collections.Hashtable properities = new System.Collections.Hashtable();
                            foreach (SPField field in file.Item.Fields)
                            {
                                if (field.Group != "_Hidden" && !field.Hidden && !field.Sealed && field.InternalName != "ID" && !field.ReadOnlyField)
                                    properities.Add(field.InternalName, file.Item[field.InternalName]);
                            }
                            foreach (SPFileVersion fileVersion in fileVersionCollection)
                            {
                                SPFile newVesionFile = folder.Files.Add(file.Item.Url, fileVersion.OpenBinaryStream(), properities, createdUser, modifiedUser, createdTime, modifiedTime, file.CheckInComment, true);
                                _log.InfoFormat("Succeeded to add file Version : {0} and the  url: {1} In Site : {2}", Convert.ToString(newVesionFile.Item["FileLeafRef"]), fileVersion.Url, targetWeb.Title);
                            }

                            SPFile newFile = folder.Files.Add(file.Item.Url, file.OpenBinaryStream(), properities, createdUser, modifiedUser, createdTime, modifiedTime, file.CheckInComment, true);
                            _log.InfoFormat("Succeeded to add file : {0} and the  url: {1} In Site : {2}", Convert.ToString(newFile.Item["FileLeafRef"]), file.Item.Url, targetWeb.Title);
                            // write the id for succeeded Migrated file
                            WriteMigratedIds(file.Item.ID, string.Empty);

                            /*
                           foreach (SPField field in file.Item.Fields)
                           {
                               if (field.Group != "_Hidden" && !field.Hidden && !field.Sealed && field.InternalName != "ID" && !field.ReadOnlyField)
                                   newFile.Item[field.InternalName] = file.Item[field.InternalName];
                           }*/
                            /*
                            newFile.Item.Update();
                            _log.InfoFormat("Succeeded to add MetaData to File: {0} and the url: {1} in Site: {2}", Convert.ToString(newFile.Item["FileLeafRef"]), newFile.Url, targetWeb.Title);
                            */

                        }

                    }
                    catch (Exception innerEx)
                    {
                        WriteFailedContenttypes(innerEx.Message);
                        _log.FatalFormat("Exception Happened when Migrate file: {0} To Site: {1} Exception : {2}", Convert.ToString(file.Item["FileLeafRef"]), targetWeb.Name, innerEx.Message);
                        WriteFailedMigratedIds(file.Item.ID, file.Item.Url);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteFailedContenttypes(ex.Message);
                _log.FatalFormat("Exception : {0}", ex.Message);
            }
        }

        /// <summary>
        /// Add SubFolders and items to Document library in target Web 
        /// </summary>
        /// <param name="folders"></param>
        /// <param name="targetWeb"></param>
        /// <param name="sourceWeb"></param>
        /// <param name="docLibraryName"></param>
        public static void GenerateFolders(SPFolderCollection folders, SPWeb targetWeb, SPWeb sourceWeb, string docLibraryName)
        {
            foreach (SPFolder folder in folders)
            {
                try
                {
                    string folderUrl = string.Empty;
                    string[] folderurlSplitted = folder.ServerRelativeUrl.Split('/');
                    for (int i = 1; i < folderurlSplitted.Length - 1; i++)
                    {
                        folderUrl += "/" + folderurlSplitted[i];
                    }
                    if (targetWeb.ServerRelativeUrl != "/")
                        folderUrl = targetWeb.ServerRelativeUrl + folderUrl;

                    bool migratedId = false;
                    if (migratedItemIds != null)
                    {
                        string[] values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(folder.Item.ID))).ToList(), folder.Item.ID);
                        migratedId = values[0] == Convert.ToString(folder.Item.ID);
                    }
                    if (!migratedId)
                    {
                        SPListItem itemCreation = targetWeb.Lists[docLibraryName].Items.Add(folderUrl, folder.Item.FileSystemObjectType, Convert.ToString(folder.Item["FileLeafRef"]));
                        itemCreation.Update();
                        _log.InfoFormat("Create folder: {0} and the Url is {1}", Convert.ToString(folder.Item["FileLeafRef"]), folderUrl);
                        foreach (SPField field in folder.Item.Fields)
                        {
                            if (field.Group != "_Hidden" && !field.Hidden && !field.Sealed && field.InternalName != "ID" && !field.ReadOnlyField)
                            {
                                itemCreation[field.InternalName] = folder.Item[field.InternalName];
                            }
                        }
                        itemCreation.Update();
                        _log.InfoFormat("Succeeded to add MetaData to Folder: {0} in Site: {1}", Convert.ToString(folder.Item["FileLeafRef"]), targetWeb.Title);
                        WriteMigratedIds(folder.Item.ID, itemCreation.Folder.Url);
                        if (folder.Files.Count > 0)
                            GenerateFiles(folder, itemCreation.Folder, targetWeb, sourceWeb);
                        if (folder.SubFolders.Count > 0)
                        {
                            GenerateFolders(folder.SubFolders, targetWeb, sourceWeb, docLibraryName);
                        }
                    }
                    else
                    {
                        string[] values = null;


                        values = GetValuesForMigratedLine(migratedItemIds.Where(c => c.Contains(Convert.ToString(folder.Item.ID))).ToList(), folder.Item.ID);
                        // migratedId = values[0] == Convert.ToString(item.ID);


                        SPFolder targetFolder = targetWeb.GetFolder(values[1]);
                        if (folder.Files.Count > 0)
                            GenerateFiles(folder, targetFolder, targetWeb, sourceWeb);
                        if (folder.SubFolders.Count > 0)
                        {
                            GenerateFolders(folder.SubFolders, targetWeb, sourceWeb, docLibraryName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.FatalFormat("Exception Happended when Migrate Folder: {0} from Site: {1} to Site: {2} Exception: {3}", folder.Name, sourceWeb.Name, targetWeb.Name, ex.Message);
                    WriteFailedMigratedIds(folder.Item.ID, folder.Url);
                }
            }
        }


        /// <summary>
        /// Get Ids for folders that we need to put it in specific site Collection
        /// </summary>
        /// <returns></returns>
        public static List<string> GetFolderIdsFromExcel()
        {
            try
            {
                List<string> folderNIds = new List<string>();
                string urlExcel = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["urlExcel"]);
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(urlExcel);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                        folderNIds.Add(Convert.ToString(xlRange.Cells[i, 1].Value2));
                }
                _log.InfoFormat("Read folder Ids from excel");
                xlWorkbook.Close();
                xlApp.Quit();
                return folderNIds;
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception: {0}", ex.Message);
                return null;
            }


        }


        /// <summary>
        /// Get Folders Id
        /// </summary>
        /// <returns></returns>
        public static string[] GetFolderIdsFromText()
        {
            string urlForFolderIds = string.Empty;
            try
            {
                urlForFolderIds = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["UrlForFoldertxt"]);
                string[] lines = File.ReadAllLines(urlForFolderIds);
                return lines;
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Can not read Folder Ids from the file, Exception: {0}", ex.Message);

                return null;
            }


        }
        /// <summary>
        /// get from migrated file the ids for the items that migrated well
        /// </summary>
        /// <returns></returns>
        public static string[] ReadMigratedIds()
        {
            string succeededMigratedFilePath = string.Empty;
            try
            {
                succeededMigratedFilePath = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SucceededMigratedFile"]);
                string[] lines = File.ReadAllLines(succeededMigratedFilePath);
                return lines;
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Can not read Migrated Ids from the file, Exception: {0}", ex.Message);
                FileStream fileStream = File.Create(succeededMigratedFilePath);
                if (fileStream != null)
                    _log.InfoFormat("Create file for suceeded migrated items in path: {0}", succeededMigratedFilePath);
                return null;
            }

        }

        /// <summary>
        /// Add Id and Url for the item That Migrated Well.
        /// </summary>
        /// <param name="Id"></param>
        /// <param name="folderUrl"></param>
        public static void WriteMigratedIds(int Id, string folderUrl)
        {
            try
            {
                string succeededMigratedFilePath = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["SucceededMigratedFile"]);
                using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(succeededMigratedFilePath, true))
                {

                    if (string.IsNullOrEmpty(folderUrl))
                    {
                        file.WriteLine(Id);
                    }
                    else
                    {
                        file.WriteLine(Id + "-" + folderUrl);
                    }
                    file.Close();
                }

            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception: {0}", ex.Message);
            }
        }

        /// <summary>
        /// Add Id and Url for the item That Failed to Migrated. 
        /// </summary>
        /// <param name="Id"></param>
        /// <param name="itemUrl"></param>
        public static void WriteFailedMigratedIds(int Id, string itemUrl)
        {
            string failedMigratedFile = string.Empty;
            try
            {
                failedMigratedFile = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["FailedMigratedFile"]);
                using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(failedMigratedFile, true))
                {
                    if (string.IsNullOrEmpty(itemUrl))
                    {
                        file.WriteLine(Id);
                    }
                    else
                    {
                        file.WriteLine(Id + "-" + itemUrl);
                    }
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception: {0}", ex.Message);
                FileStream fileStream = File.Create(failedMigratedFile);
                if (fileStream != null)
                    _log.InfoFormat("Create file for suceeded migrated items in path: {0}", failedMigratedFile);
                WriteFailedMigratedIds(Id, itemUrl);
            }
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="txt"></param>
        public static void WriteFailedContenttypes(string txt)
        {
            string failedContenttypes = string.Empty;
            try
            {
                failedContenttypes = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["FailedContenttypes"]);
                using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(failedContenttypes, true))
                {
                    file.WriteLine(txt);
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception: {0}", ex.Message);
                FileStream fileStream = File.Create(failedContenttypes);
                if (fileStream != null)
                    _log.InfoFormat("Create file for failed Content types in path: {0}", failedContenttypes);
                WriteFailedContenttypes(txt);
            }
        }

        /// <summary>
        /// Get line and their value that correspending to itemID  and return their values(Id & url).
        /// </summary>
        /// <param name="collectionOfString"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        public static string[] GetValuesForMigratedLine(List<string> collectionOfString, int itemId)
        {
            string[] retValues = new string[2];
            try
            {
                foreach (string line in collectionOfString)
                {

                    if (line.Contains('-'))
                    {
                        string[] values = line.Split('-');
                        if (Convert.ToInt64(values[0]) == itemId)
                        {
                            retValues[0] = values[0];
                            retValues[1] = values[1];
                            break;
                        }
                    }
                    else
                    {
                        if (Convert.ToInt64(line) == itemId)
                        {
                            retValues[0] = line;
                            break;
                        }
                    }
                }
                return retValues;
            }
            catch (Exception ex)
            {
                _log.FatalFormat("Exception: {0}", ex.Message);
                return null;
            }
        }

    }
}
