using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ComboxManager.Business.Model;
using ComboxManager.Business.Model.Entities;
using ComboxManager.Model;
using JMSL.Framework.Business;
using JMSL.Framework.Business.Entities;
using JMSL.Framework.Business.Services;
using JMSL.Framework.Log;
using JMSL.Framework.Utility;

namespace DocumentImport
{
    public class ProjectDocumentImporter
    {
        #region Private members
        private string _rootFolder;
        private string _templateFolder;
        private double totalSize = 0; // in mb
        private List<string> _excludedEndFiles = new List<string>() {
            "Thumbs.db", "-NC.txt", "toto.zip", "toto1.zip"
        };
        private List<string> _excludedFileExtensions = new List<string>() {
                        
            "tmp", "old", "old2", "old3", "lnk", "err", "old1", "", "015", "300", "scrap",

            //Catia extension
            //"CATPart", "CatProduct", "catdrawing", "debug", "original", "tools-6", "tools-6bup", "stp","MIN", "tls", "NC", "log", "dxf",  "xlo",  
            
            //Vericut extensions
            "Stl", "ply", "vcproject", "swp", "lib", "min", "mch", "ctl", "ssb", "ops", "vctmp", "stl", "lwn", "ip", "sat", "vcproject_ext", "VcInspect",

            // MasterCam
            "MCX-7", "MCX-6", "MCX","mcx-8", "mc9", "igs", "tooldb", "mcx-5", "lck",
        };
        private Dictionary<string, string> _projectProgram = new Dictionary<string, string>()
        {
            {"113W2155-33", "Commercial"},
            {"113W2155-35", "Commercial"},
            {"141W6314-15", "Commercial"},
            {"141W6314-16", "Commercial"},
            {"141W6315-4", "Commercial"},
            {"150-621-1", "Commercial"},
            {"161W2201-1", "Commercial"},
            {"161W2305-1", "Commercial"},
            {"200-0488-103", "Militaire"},
            {"200-2321-101", "Militaire"},
            {"3002680-3LP", "Militaire"},
            {"337094", "Militaire"},
            {"337094-L", "Militaire"},
            {"465Z5303-151", "Commercial"},
            {"465Z5303-152", "Commercial"},
            {"465Z5303-153", "Commercial"},
            {"465Z5303-156", "Commercial"},
            {"465Z5303-161", "Commercial"},
            {"465Z5303-163", "Commercial"},
            {"465Z5303-165", "Commercial"},
            {"465Z5303-151PF", "Commercial"},
            {"465Z5303-152PF", "Commercial"},
            {"465Z5303-153PF", "Commercial"},
            {"465Z5303-156PF", "Commercial"},
            {"465Z5303-161PF", "Commercial"},
            {"465Z5303-163PF", "Commercial"},
            {"465Z5303-165PF", "Commercial"},
            {"55-1209035-01", "Commercial"},
            {"55-2397003-00", "Commercial"},
            {"69-4263", "Militaire"},
        };

        private Dictionary<string, string> _folderMapping = new Dictionary<string, string>()
        {
            {"Purchase order".ToUpper(), @"Achat\Archive"},
            {"Purchase orders".ToUpper(), @"Achat\Archive"},
            {"Customer purchase order".ToUpper(), @"Achat\Archive\Bon de commande client"},
            {"Customer request for quotes".ToUpper(), @"Achat\Archive\Demande de soumission client"},
            {"Contract review".ToUpper(), @"Méthode\Revue de contrat"},
            {"CONTRAC REVIEW".ToUpper(), @"Méthode\Revue de contrat"},
            {"Gamme".ToUpper(), @"Méthode\Gamme"},
            {"MCR".ToUpper(), @"Méthode\MCR"},
            {"Released".ToUpper(), @"Méthode\Document approuvé par le client"},
            {"Serials".ToUpper(), @"Méthode\Contrôle des numéros de série"},
            {"Approvals".ToUpper(), @"Méthode\Approbation client"},
            {"Tech & approval".ToUpper(), @"Méthode\Approbation client"},
            {"TECHNIQUE & APPROVAL".ToUpper(), @"Méthode\Approbation client"},
            {"TECHNIQUE AND APPROVAL".ToUpper(), @"Méthode\Approbation client"},
            {"techniques&approvals".ToUpper(), @"Méthode\Approbation client"},
            {"TECH-APPROVALS".ToUpper(), @"Méthode\Approbation client"},
            {"Communication".ToUpper(), @"Méthode\Communication"},
            {"COMUNICATION HEROUX".ToUpper(), @"Méthode\Communication"},
            {"COMMUNICATION HEROUX".ToUpper(), @"Méthode\Communication"},
            {"COMMUNICATION CLIENT".ToUpper(), @"Méthode\Communication"},
            {"COMMUNICATION EMAIL".ToUpper(), @"Méthode\Communication"},
            {"COMMUNICATION HD".ToUpper(), @"Méthode\Communication"},
            {"Work order".ToUpper(), @"Qualité\Bon de production"},
            {"Workorder".ToUpper(), @"Qualité\Bon de production"},
            {"Work orders".ToUpper(), @"Qualité\Bon de production"},
            {"FAI".ToUpper(), @"Qualité\FAI"},
            {"F.A.I".ToUpper(), @"Qualité\FAI"},
            {"NCR".ToUpper(), @"Qualité\Non-Conformité"},
            {"non conformity request".ToUpper(), @"Qualité\Non-Conformité"},
            {"NCR'S".ToUpper(), @"Qualité\Non-Conformité"},
            {"NCR FOURNISSEUR".ToUpper(), @"Qualité\Non-Conformité"},
            {"SNCR".ToUpper(), @"Qualité\Non-Conformité"},
            {"Croquis pdf".ToUpper(), @"Méthode\Archive"},
            {"CLIENT QUOTES AND SUBMISIONS".ToUpper(), @"Ventes\Archive"},
            {"SFTECH QUOTES FOR PROCESSING AND MATERIAL".ToUpper(), @"Achat\Archive"},
            {"Quote".ToUpper(), @"Ventes\Archive"},
            {"Quotes".ToUpper(), @"Ventes\Archive"},
            {"QUOTES FOR PROCESSING AND MATERIALS".ToUpper(), @"Ventes\Archive"},

        };
        #endregion

        public ProjectDocumentImporter(string folder, string templateFolder)
        {
            this._rootFolder = folder;
            this._templateFolder = templateFolder;
            ProcessFolder(folder, null, string.Empty);
            CLLogger.LogVerbose(string.Format("The total size of the folder '{0}' is '{1}' mb", folder, totalSize));
        }

        private void ProcessFolder(string folder, Guid? parentProjectId, string comboxPath)
        {
            if (Directory.Exists(folder))
            {

                foreach (var folderInfo in new DirectoryInfo(folder).GetDirectories())
                {
                    var projectId = Cache.GetProjectId(folderInfo.Name, folderInfo.FullName);
                    var originalComboxPath = comboxPath;
                    if (projectId == null)
                    {
                        projectId = parentProjectId;
                    }

                    if (parentProjectId == null || projectId != parentProjectId)
                    {
                        comboxPath = this._templateFolder + @"\";
                    }
                    if (projectId == Guid.Empty)
                    {
                        //CLLogger.LogVerbose(string.Format("The folder is excluded: '{0}'", folderInfo.FullName));
                    }
                    else if (projectId != null)
                    {/*
                        var acceptedProject = new List<string>() {                             
                            "1006M3205C001",
                            "1006M3300C001"
                             };
                        if (acceptedProject.Contains(Cache.GetProjectName(projectId.Value)))
                        {
                            if (_projectProgram.ContainsKey(folderInfo.Name))
                            {
                                UpdateProjectProgram(projectId.Value, _projectProgram[folderInfo.Name]);
                            }
                            else if (_projectProgram.ContainsKey(Cache.GetProjectName(projectId.Value)))
                            {
                                UpdateProjectProgram(projectId.Value, _projectProgram[Cache.GetProjectName(projectId.Value)]);

                                if (!folderInfo.Name.ToUpper().Contains("COMMERCIAL") && !folderInfo.Name.ToUpper().Contains("MILITAIRE"))
                                {
                                    comboxPath += folderInfo.Name + @"\";
                                }
                            }
                            else if (folderInfo.Name.ToUpper().Contains("COMMERCIAL"))
                            {
                                UpdateProjectProgram(projectId.Value, "Commercial");
                            }
                            else if (folderInfo.Name.ToUpper().Contains("MILITAIRE"))
                            {
                                UpdateProjectProgram(projectId.Value, "Militaire");
                            }
                            else if (projectId == parentProjectId)
                            {
                                comboxPath += folderInfo.Name + @"\";
                            }

                            if (_folderMapping.ContainsKey(folderInfo.Name.ToUpper()))
                            {
                                comboxPath = _folderMapping[folderInfo.Name.ToUpper()] + @"\";
                            }

                            ProcessFolder(folderInfo.FullName, projectId, comboxPath);

                            foreach (var fileInfo in folderInfo.GetFiles())
                            {
                                if (!fileInfo.Name.StartsWith("~") && !_excludedEndFiles.Any(x => fileInfo.Name.Trim().EndsWith(x, StringComparison.InvariantCultureIgnoreCase)))
                                {
                                    if (!_excludedFileExtensions.Any(x => x.Equals(fileInfo.Extension.Trim(), StringComparison.InvariantCultureIgnoreCase) ||
                                                                          ("." + x).Equals(fileInfo.Extension.Trim(), StringComparison.InvariantCultureIgnoreCase)))
                                    {
                                        ProcessFile(fileInfo, projectId.Value, comboxPath);
                                    }
                                    else
                                    {
                                        CLLogger.LogVerbose(string.Format("The file '{0}' with extension '{1}' will be excluded for project '{2}'", fileInfo.FullName, fileInfo.Extension, Cache.GetProjectName(projectId.Value)));
                                    }
                                }
                                else
                                {
                                    CLLogger.LogVerbose(string.Format("The file '{0}' will be excluded for project '{1}'", fileInfo.FullName, Cache.GetProjectName(projectId.Value)));
                                }
                            }
                        }
                        */
                    }
                    else
                    {
                        CLLogger.LogVerbose(string.Format("The folder was not imported: '{0}'", folderInfo.FullName));
                        //ProcessFolder(folderInfo.FullName, null, string.Empty);
                    }
                    comboxPath = originalComboxPath;
                }
            }
            else
            {
                CLLogger.LogVerbose(string.Format("The following path was not found : '{0}'", folder));
            }
        }

        private void ProcessFile(FileInfo file, Guid projectId, string comboxPath)
        {
            CLLogger.LogVerbose(string.Format("The file '{0}' --> '{2}' will be associated to project '{1}'", file.FullName.Substring(_rootFolder.Length), Cache.GetProjectName(projectId), comboxPath));
            totalSize += (file.Length * 1d) / 1024 / 1024;

            var cleanPath = comboxPath.TrimEnd('\\').Split('\\');
            Guid? folderId = null;

            using (var context = FactoryDbContext.Create())
            {
                for (var i = 0; i < cleanPath.Length; i++)
                {
                    var currentFolder = CleanStringForWeb(cleanPath[i]);
                    var existingFolderId = context.EntityFolders.Where(x => x.ReferenceId == projectId && x.ParentId == folderId && x.Name == currentFolder).Select(x => x.Id).FirstOrDefault();

                    if (existingFolderId == Guid.Empty)
                    {
                        var folder = EntityFolder.New();
                        folder.CurrentUserId = Constants.User.Admin;
                        folder.Id = Guid.NewGuid();
                        folder.EntityId = Constants.EntityType.Project;
                        folder.ReferenceId = projectId;
                        folder.ParentId = folderId;
                        folder.Name = currentFolder;
                        folder.Save();

                        folderId = folder.Id;
                    }
                    else
                    {
                        folderId = existingFolderId;
                    }
                }
            }

            var document = Document.New();

            document.ReferenceId = projectId;
            document.EntityId = Constants.EntityType.Project;
            document.CurrentUserId = Constants.User.Admin;
            document.CreatedOn = DateTime.Now;
            document.LastUpdatedOn = DateTime.Now;
            document.CreatedById = Constants.User.Admin;
            document.LastUpdatedById = Constants.User.Admin;
            document.Id = Guid.NewGuid();
            document.FolderId = folderId;
            document.DocumentTypeId = LookupServices.GetLookupItemIdFromTag("LST_DOCUMENT_TYPE_NOT_CLASSIFIED");
            document.Filename = CleanStringForWeb(file.Name).Replace("..", ".");
            document.FileExtension = file.Extension;
            document.FileContentType = ContentTypes.GetContentTypeFromFileExtension(document.FileExtension);

            try
            {
                using (var stream = file.OpenRead())
                {
                    document.Insert<JMSL.Framework.DAL.Entities.Document, Document>(DocumentManagerProvider.Get(),
                                                                                    new TransferDocument()
                                                                                    {
                                                                                        FileName = document.Filename,
                                                                                        FileLength = file.Length,
                                                                                        FileByteStream = stream
                                                                                    });
                }
            }
            catch (Exception ex)
            {
                CLLogger.LogError(ex);
            }
        }

        private byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        private string CleanStringForWeb(string name)
        {
            var excludedChar = "#%&*:<>?/{|}";
            excludedChar.ToCharArray().ToList().ForEach(x => name = name.Replace(x.ToString(), ""));
            return name;
        }

        private void UpdateProjectProgram(Guid projectId, string programName)
        {
            var projectNeedsToBeUpdated = false;
            var programId = Cache.GetProjectProgramId(programName);
            using (var context = FactoryDbContext.Create())
            {
                projectNeedsToBeUpdated = context.Projects.Any(x => x.Id == projectId && x.ProgramId != programId);
            }

            if (projectNeedsToBeUpdated)
            {
                var project = Project.Get(projectId);

                project.ProgramId = programId;
                project.CurrentUserId = Constants.User.Admin;
                project.IsModifiedByImport = true;
                project.Save();
                CLLogger.LogVerbose(string.Format("The project '{0}' program as been updated to '{1}'", project.Name, programName));
                project = null;
            }
        }
    }
}






