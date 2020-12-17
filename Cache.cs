using System;
using System.Collections.Generic;
using System.Linq;
using ComboxManager.Model;
using JMSL.Framework.DAL.Entities;

namespace DocumentImport
{
    public static class Cache
    {
        private static Dictionary<string, Guid> _projects = null;
        private static Dictionary<string, Guid> _programs = null;

        #region Replace Part of the path

        private static Dictionary<string, string> _replaceEnd = new Dictionary<string, string>()
        {
            {"_A", ""}, //1003m3602c001, 293w3363-1
            {"LP", ""}, //200-0301-103, 200-0302-101, 200-1321-101
            {" ISTCR", ""}, //4118-0001, 4118-0006, 4118-0011, 4118-0012, 4125-0001, 4125-0004, 4125-0201
            {"-L", ""}, //337094, 371689-5
            {"PF", ""}, // 465Z5303-151, 465Z5303-152, 465Z5303-153, 465Z5303-155, 465Z5303-156, 465Z5303-161, 465Z5303-163, 465Z5303-165
            {" -test 12po", "-12\"-TESTPIECE"},
            {"-test 8po", "-8\"-TESTPIECE"},
            {"-1 ASO", "-1LP"}, //435564-1 ASO
            {"-11 LP", "-11LP"}, //352604-11 LP
            {"371689 #2 TERMINAL FORGING", "371689-2"},
            {"371689-3 TERMINAL#3", "371689-3"},
            {"371689_12 INCH TEST PC MACH", "371689-12\"-TESTPIECE"},
            {"371689_8 INCH TEST PC MACH", "371689-8\"-TESTPIECE"},
            {"800905-109-ASO", "800905"},
            {"801445-3", "801445"},
            {"802290-1", "802290"},
            {"805774-101", "805774"},
            {"901543-1", "901543"},
            {"901678 TP1-2-3-4", "901678-TP4"},
            {"902387-3LP", "902387"},
            {"DAA3221A154-101", "DAA3221A154-001"},
        };

        #endregion
        #region Excluded Path
        private static List<string> _excludedPath = new List<string>() 
        {
            @"P:\Ingenierie\PROGRAMMATION\ACHAT",
            @"P:\Ingenierie\PROGRAMMATION\CH47 Material",
            @"P:\Ingenierie\PROGRAMMATION\Cheminee lining",
            @"P:\Ingenierie\PROGRAMMATION\CNC Machine Network",
            @"P:\Ingenierie\PROGRAMMATION\cnc_machines",
            @"P:\Ingenierie\PROGRAMMATION\driver  usb232r-10-bulk",
            @"P:\Ingenierie\PROGRAMMATION\FacePlate",
            @"P:\Ingenierie\PROGRAMMATION\Fiftih Axis tests",
            @"P:\Ingenierie\PROGRAMMATION\fixture rti mori",
            @"P:\Ingenierie\PROGRAMMATION\INFORMATION DIVERS",
            @"P:\Ingenierie\PROGRAMMATION\JSR",
            @"P:\Ingenierie\PROGRAMMATION\lathe",
            @"P:\Ingenierie\PROGRAMMATION\LIBRAIRIE MASTERCAM _ VINCENT",
            @"P:\Ingenierie\PROGRAMMATION\Logiciel",
            @"P:\Ingenierie\PROGRAMMATION\Mastercam",
            @"P:\Ingenierie\PROGRAMMATION\Mastercam Backup",
            @"P:\Ingenierie\PROGRAMMATION\Messier",
            @"P:\Ingenierie\PROGRAMMATION\mill",
            @"P:\Ingenierie\PROGRAMMATION\MISC FILES",
            @"P:\Ingenierie\PROGRAMMATION\MoldPlus",
            @"P:\Ingenierie\PROGRAMMATION\New folder",
            @"P:\Ingenierie\PROGRAMMATION\part-number",
            @"P:\Ingenierie\PROGRAMMATION\posts",
            @"P:\Ingenierie\PROGRAMMATION\Projets",
            @"P:\Ingenierie\PROGRAMMATION\RecycleBin",
            @"P:\Ingenierie\PROGRAMMATION\Simulation",
            @"P:\Ingenierie\PROGRAMMATION\Software",
            @"P:\Ingenierie\PROGRAMMATION\soumission",
            @"P:\Ingenierie\PROGRAMMATION\Temp",
            @"P:\Ingenierie\PROGRAMMATION\Templates",
            @"P:\Ingenierie\PROGRAMMATION\Test RTI Claro",
            @"P:\Ingenierie\PROGRAMMATION\TEST SWAGGING",
            @"P:\Ingenierie\PROGRAMMATION\Tooling",
            @"P:\Ingenierie\PROGRAMMATION\Tutorials and tips",
            @"P:\Ingenierie\PROGRAMMATION\Vericut",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\CLI_Pieces Remplacement",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\HD LAVAL FIXTURE",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\Mecaer G650 Side Brace Measurements",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\P3 Gamme",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\Techfab - Inspection",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\temp",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\tempDANNY PAQUETTE FEB 17_2014",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\TM-LTS-2",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\A109 Casting",
            @"P:\Ingenierie\PART DOCUMENTS\PRODUCTION PART\B777",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
            @"",
        };
        #endregion

        public static Guid? GetProjectId(string projectNumber, string fullPath)
        {
            if (_excludedPath.Any(x => x.Equals(fullPath, StringComparison.InvariantCultureIgnoreCase)))
            {
                return Guid.Empty;
            }
            if (_projects == null)
            {
                using (var context = FactoryDbContext.Create())
                {
                    _projects = context.Projects.Select(x => new { x.Id, x.ProjectNumber }).ToDictionary(x => x.ProjectNumber.Trim(), x => x.Id);
                }
            }

            if (_projects.ContainsKey(projectNumber))
            {
                return _projects[projectNumber];
            }

            foreach (var x in _replaceEnd)
            {
                var match = _projects.Where(p => ReplaceEnd(p.Key, x).Equals(ReplaceEnd(projectNumber, x), StringComparison.InvariantCultureIgnoreCase)).Select(p => p.Value).FirstOrDefault();

                if (match != Guid.Empty)
                {
                    return match;
                }
            };


            return null;
        }

        public static string GetProjectName(Guid projectId)
        {
            return _projects.Where(x => x.Value == projectId).Select(x => x.Key).FirstOrDefault();
        }

        public static Guid GetProjectProgramId(string programName)
        {
            if (_programs == null)
            {
                _programs = new Dictionary<string, Guid>();
                using (var context = FactoryDbContext.Create())
                {
                    var program = context.Lookups.Include("Items.Descriptions").Where(x => x.Tag == "LST_PROJECT_PROGRAM").FirstOrDefault();

                    if (program != null)
                    {
                        _programs = program.Items.Select(x => new { x.Id, x.Descriptions.First(d => d.CultureId == Constants.Culture.Francais).Text }).ToDictionary(x => x.Text, x => x.Id);
                    }
                }
            }

            return _programs.First(x => x.Key.ToUpper().Contains(programName.ToUpper())).Value;
        }
        private static string ReplaceEnd(string originalString, KeyValuePair<string, string> toReplace)
        {
            if (originalString.EndsWith(toReplace.Key))
            {
                return originalString.Substring(0, originalString.LastIndexOf(toReplace.Key)) + toReplace.Value;
            }
            return originalString;
        }
    }
}
