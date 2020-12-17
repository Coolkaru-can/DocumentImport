using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using ComboxManager.Model;
using JMSL.Framework.Business;
using JMSL.Framework.Extensions;
using JMSL.Framework.Mail;

namespace DocumentImport
{
    public static class RenameEntityFolderSharepointProcess
    {
        public static void Execute()
        {
            var manager = DocumentManagerProvider.Get() as SharepointDocumentManager;

            if (manager != null && manager.SharepointUseInternalTagFolder)
            {
                using (var db = FactoryDbContext.Create())
                {
                    var sql = @"WITH Folders (Id, Name, FullPath, EntityId, ReferenceId, InternalTag, IsProcessed, Level)
                                AS
                                (
                                    SELECT F.Id, F.Name, F.FullPath, EF.EntityId, EF.ReferenceId, CAST(InternalId AS NVARCHAR(MAX)) AS InternalTag, EF.IsProcessed, 0 AS Level
                                    FROM Folder F
                                    INNER JOIN EntityFolder EF ON F.Id = EF.Id
                                    WHERE F.ParentId IS NULL
                                    UNION ALL
                                    SELECT F.Id, F.Name, F.FullPath, EF.EntityId, EF.ReferenceId, CAST(F.InternalId AS NVARCHAR(MAX)) AS InternalTag, EF.IsProcessed, P.Level + 1
                                    FROM Folder F
                                    INNER JOIN Folders P ON F.ParentId = P.Id
                                    INNER JOIN EntityFolder EF ON F.Id = EF.Id
                                )
                                SELECT Id, Name, FullPath, EntityId, ReferenceId, InternalTag AS Tag, IsProcessed
                                FROM Folders
                                ORDER BY EntityId, ReferenceId, Level DESC;";

                    var folders = db.Database.SqlQuery<FolderImport>(sql).ToList().Where(x => !x.IsProcessed).ToList();
                    var errors = new List<string>();
                    var i = 0;
                    var count = folders.Count;

                    folders.ForEach(folder =>
                    {
                        try
                        {
                            manager.RenameFolder(folder.EntityId, folder.ReferenceId, folder.Tag, folder.Name, folder.FullPath.HasValue() ? folder.FullPath : folder.Name, db);
                            db.Database.ExecuteSqlCommand("UPDATE EntityFolder SET IsProcessed = 1 WHERE Id = {0}", folder.Id);
                            i++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add(string.Format("Id: {0}; Name: {1}; FullPath: {2}", folder.Id, folder.Name, folder.FullPath) + Environment.NewLine + ex.Message);

                            Console.WriteLine(ex.Message);
                        }
                    });

                    if (errors.HasValue())
                    {
                        var message = errors.JoinBy(Environment.NewLine);

                        MailUtilities.SendMail(new List<string>() { ConfigurationManager.AppSettings["AlertRecipient"] }, "ERROR Rename Folder SharePoint", message);
                    }
                }
            }
        }

        private class FolderImport
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public string FullPath { get; set; }
            public Guid EntityId { get; set; }
            public Guid ReferenceId { get; set; }
            public string Tag { get; set; }
            public bool IsProcessed { get; set; }
        }
    }
}
