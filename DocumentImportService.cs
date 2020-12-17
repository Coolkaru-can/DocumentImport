using System;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using ComboxManager.Business.Extensions;
using ComboxManager.Business.Services;
using ComboxManager.Model;
using JMSL.Framework.Business;
using JMSL.Framework.Business.Interfaces;
using JMSL.Framework.Extensions;
using JMSL.Framework.Log;

namespace DocumentImport
{
    public partial class DocumentImportService : ServiceBase
    {
        public DocumentImportService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Start();
        }

        public static void Start()
        {
            new Timer(new TimerCallback(Synchronize), null, 100, -1);
        }

        public void Synchronize()
        {
            Synchronize(null);
        }

        protected static void Synchronize(object stateInfo)
        {
            try
            {
                var startingTime = DateTime.Now;
                SetDefaultDocumentManager();
                //new ProjectDocumentImporter(@"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\", @"Amélioration continue\À classer");

                RenameEntityFolderSharepointProcess.Execute();

                CLLogger.LogVerbose("The upload takes " + DateTime.Now.Subtract(startingTime).TotalMinutes.ToString() + " minutes.");
            }
            catch (Exception ex)
            {
                CLLogger.LogCritical(ex);
            }
        }

        private static void SetDefaultDocumentManager()
        {
            App.Cache.Companies(); // load companies before in order to avoid an error of transaction scope
            var temp = App.Cache.EntityLists; // To load entities in cache

            DocumentManagerProvider.DocumentManagerChooser = new Func<IDocumentManager>(() =>
            {
                try
                {
                    var company = App.Cache.Companies().FirstOrDefault(x => x.Id == new Guid("2F018D31-A8E4-47AC-8A29-C78190768147"));

                    if (company != null && company.UseSharepoint)
                    {
                        return new SharepointDocumentManager()
                        {
                            SharepointDomain = company.SharepointDomain,
                            SharepointLibraryName = company.SharepointLibraryName,
                            SharepointUrl = company.SharepointUrl,
                            SharepointUserName = company.SharepointUserName,
                            SharepointPassword = company.SharepointPassword,
                            SharepointInCloud = company.SharepointInCloud,
                            SharepointUseInternalTagFolder = true //company.SharepointUseInternalTagFolder
                        };
                    }
                }
                catch (Exception ex)
                {
                    CLLogger.LogError(ex);
                }

                return new JMSLDocumentManager();
            });

            DocumentManagerProvider.DocumentBaseFolderName = new Func<JMSL.Framework.DAL.Entities.BaseEntitiesDbContext, JMSL.Framework.Business.Entities.EntityInfo, Guid, string>((db, entityInfo, referenceId) =>
            {
                if (entityInfo != null)
                {
                    if (entityInfo.Name.Compare("Project"))
                        return entityInfo.Name + "\\" + ProjectServices.GetProjectNumberById(db as ComboxDbContext, referenceId).Replace("\"", "");
                    else if (entityInfo.Name.Compare("Client"))
                        return entityInfo.Name + "\\" + ClientServices.GetClientInfo(db as ComboxDbContext, referenceId).Name.Replace("\"", "");
                    else if (entityInfo.Name.Compare("Documentation"))
                        return entityInfo.Name;
                    else
                        return entityInfo.Name + "\\";
                }

                return string.Empty;
            });
        }
    }
}
