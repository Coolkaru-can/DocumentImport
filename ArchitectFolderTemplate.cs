using ComboxManager.Business.Model;
using ComboxManager.Model.Entities;
using DSOFile;
using JMSL.Framework.Log;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocumentImport
{
    public static class ArchitectFolderTemplate
    {
        public static void ImportFile()
        {
            ChangeFileCustomProperties();
        }

        private static FolderTemplate ImportDivision(Guid folderTemplateHeaderId, string division)
        {
            CLLogger.LogInformation(string.Format(@"Found division '{0}'", division));
            return CreateNewFolderTemplate(folderTemplateHeaderId, null, division);
        }
        private static FolderTemplate ImportFolder(Guid folderTemplateHeaderId, Guid folderTemplateParentId, string division, string folder)
        {
            var date = folder.Trim().Substring(folder.Trim().Length - 10);
            var cleanFolderName = folder.Trim().Substring(0, folder.Trim().Length - 10).Trim();

            CLLogger.LogInformation(string.Format(@"Found folder '{0}\{1}' Date : '{2}'", division, cleanFolderName, date));
            return CreateNewFolderTemplate(folderTemplateHeaderId, folderTemplateParentId, cleanFolderName);
        }

        private static string GetFileContent()
        {
            var file = new System.IO.FileStream(@"C:\Temp\Architect.txt", FileMode.Open);
            var sr = new StreamReader(file);

            file.Position = 0;
            var fileContent = sr.ReadToEnd();
            file.Close();
            sr.Close();
            
            return fileContent;
        }

        private static void ChangeFileCustomProperties()
        {
            var doc = new OleDocumentPropertiesClass();

            try
            {
                doc.Open(@"C:\test danny.txt");
            //doc.SummaryProperties.Company = "ComboxTest";
            doc.CustomProperties.Add("ComboxManager", Guid.NewGuid().ToString());
            }            
            catch(Exception ex)
            {
                CLLogger.LogError(ex);
                Console.WriteLine(ex.Message);
                ex = null;
            }

            //after making changes, you need to use this line to save them
            doc.Save();
        }

        private static bool EndsWithDate(string s)
        {
            DateTime d;
            return s.Trim().Length > 10 && DateTime.TryParse(s.Trim().Substring(s.Trim().Length - 10), out d);
        }

        private static FolderTemplateHeader CreateNewFolderTemplateHeader()
        {
            return new FolderTemplateHeader()
            {
                CompanyId = new Guid("2F018D31-A8E4-47AC-8A29-C78190768147"),
                CreatedById = Constants.User.Admin,
                CreatedOn = DateTime.Now,
                EntityId = Constants.EntityType.Project,
                Id = Guid.NewGuid(),
                LastUpdatedById = Constants.User.Admin,
                LastUpdatedOn = DateTime.Now,
                FolderTemplates = new List<FolderTemplate>()
            };
        }
        private static FolderTemplate CreateNewFolderTemplate(Guid folderTemplateHeaderId, Guid? parentFolderTemplateId, string folderName)
        {
            var localizationId = Guid.NewGuid();
            return new FolderTemplate()
            {
                FolderTemplateHeaderId = folderTemplateHeaderId,
                Id = Guid.NewGuid(),
                Name = folderName,
                ParentId = parentFolderTemplateId,
                FolderTemplates = new List<FolderTemplate>(),
                FolderTemplateSecurities = new List<FolderTemplateSecurity>()
                {
                    new FolderTemplateSecurity() {
                        CreatedById = Constants.User.Admin,
                        CreatedOn = DateTime.Now,
                        GroupId = new Guid("EE13F84A-01E7-4072-8408-6B6E7A35404F"),
                        Id = Guid.NewGuid(),
                        LastUpdatedById = Constants.User.Admin,
                        LastUpdatedOn = DateTime.Now,
                        SecurityType = 2
                    }
                }
            };
        }
    }
}
