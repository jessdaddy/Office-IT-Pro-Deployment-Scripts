using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator.Implementation
{
    public class OfficeInstallProperties : IOfficeInstallProperties
    {
        public string ExecutablePath { get; set; }

        public OfficeVersion OfficeVersion { get; set; }

        public string ConfigurationXmlPath { get; set; }

        public bool UseExternalSource { get; set; }

        public string SourceFilePath { get; set; }

        public string SPDesignerSource { get; set; }

        public string InfoPathSource { get; set; }




    }
}
