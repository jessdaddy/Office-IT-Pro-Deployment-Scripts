using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeInstallGenerator;

namespace Microsoft.OfficeProPlus.InstallGenerator
{
    public interface IOfficeInstallProperties
    {
        OfficeVersion OfficeVersion { get; set; }

        string ConfigurationXmlPath { get; set; }

        string SourceFilePath { get; set; }

        bool UseExternalSource { get; set; }

        string SPDesignerSource { get; set; }

        string InfoPathSource { get; set; }

        string ExecutablePath { get; set; }

    }
}
