using System;
using System.Configuration;
using System.Linq;
using System.ServiceProcess;
using System.Windows.Forms;
using ComboxManager.Model;
using JMSL.Framework.Extensions;
using JMSL.Framework.Log;
using JMSL.Framework.Mail;
using JMSL.Framework.Unity;
using Microsoft.Practices.Prism.Modularity;
using Microsoft.Practices.Unity;

namespace DocumentImport
{
    static class Program
    {
        public static string ServiceName { get { return JMSL.Framework.Divers.My.AppName(true); } }

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                args[0] = args[0].ToLower();
            }

            AppDomain.CurrentDomain.UnhandledException += CLLogger.AppDomainUnhandledExceptionHandler;
            Application.ThreadException += CLLogger.ThreadExceptionHandler;

            MailUtilities.DefaultSmtpServer = ConfigurationManager.AppSettings["DefaultSmtpServer"];
            MailUtilities.DefaultFromAdress = ConfigurationManager.AppSettings["DefaultSmtpFromAdress"];
            MailUtilities.DefaultUserName = ConfigurationManager.AppSettings["DefaultSmtpUserName"];
            MailUtilities.DefaultPassword = ConfigurationManager.AppSettings["DefaultSmtpPassword"];

            if (ConfigurationManager.AppSettings["DefaultErrorEmailRecipient"].HasValue())
                CLLogger.DefaultErrorEmailRecipient.Add(ConfigurationManager.AppSettings["DefaultErrorEmailRecipient"]);
            else if (ConfigurationManager.AppSettings["AlertRecipient"].HasValue())
                CLLogger.DefaultErrorEmailRecipient.Add(ConfigurationManager.AppSettings["AlertRecipient"]);

            CLLogger.LogInformation(String.Format("Started with args.count = {0} Args: {1}", args.Count(), String.Join(" ", args)));

            if (args.Count() == 0)
            {
                Init();
                ServiceBase[] servicesToRun = new ServiceBase[] { new DocumentImportService() };
                ServiceBase.Run(servicesToRun);
            }
            else
            {
                #region Monitor

                if (args[0] == "monitor")
                {
                    Init();
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Monitor());
                }

                #endregion

                #region Started with args... as a service controler/manager

                else
                {
                    using (var sc = new ServiceController(Program.ServiceName))
                    {
                        switch (args[0])
                        {
                            case "start":
                                sc.Start();
                                break;
                            case "stop":
                                sc.Stop();
                                break;
                            case "pause":
                                sc.Pause();
                                break;
                            case "continue":
                                sc.Continue();
                                break;
                            case "status":
                                CLLogger.LogInformation("Service status:" + sc.Status);
                                break;
                            case "install":
                                var installProcess = new System.Diagnostics.Process();
                                installProcess.StartInfo.FileName = JMSL.Framework.Divers.My.AppName() + "_Install.bat";
                                var installPath = @"c:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe /ShowCallStack " + JMSL.Framework.Divers.My.AppName(false);
                                CLLogger.LogInformation(string.Format("If you need to reinstall the service, you should run the following command or run {0}_Install.bat", JMSL.Framework.Divers.My.AppName()));
                                CLLogger.LogInformation(installPath);
                                JMSL.Framework.Divers.FileIO.WriteToFile(installProcess.StartInfo.FileName, installPath + "\r\n@pause");
                                installProcess.Start();
                                break;
                            case "uninstall":
                                var uninstallProcess = new System.Diagnostics.Process();
                                uninstallProcess.StartInfo.FileName = JMSL.Framework.Divers.My.AppName() + "_UnInstall.bat";
                                var uninstallPath = @"c:\Windows\Microsoft.NET\Framework\v4.0.30319\installUtil.exe /u /ShowCallStack " + JMSL.Framework.Divers.My.AppName(false);
                                CLLogger.LogInformation(string.Format("If you need to uninstall the service, you should run the following command or run {0}_uninstall.bat", JMSL.Framework.Divers.My.AppName()));
                                CLLogger.LogInformation(uninstallPath);
                                JMSL.Framework.Divers.FileIO.WriteToFile(uninstallProcess.StartInfo.FileName, uninstallPath + "\r\n@pause");
                                uninstallProcess.Start();
                                break;
                            default:
                                CLLogger.LogInformation("Unknown arguments:" + args[0]);
                                break;
                        }
                    }
                    return;
                }

                #endregion
            }
        }

        private static void Init()
        {
            var container = UnityContainerProvider.Initialize();
            ModuleLoader.Initialize(container);
            var catalog = new ModuleCatalog();
            container.RegisterInstance<IModuleCatalog>(catalog, new ContainerControlledLifetimeManager());
            ModuleLoader.Run(container);

            FactoryDbContext.InitDatabase(false);
        }
    }
}
