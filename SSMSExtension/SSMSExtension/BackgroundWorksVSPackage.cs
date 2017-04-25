//------------------------------------------------------------------------------
// <copyright file="BackgroundVSPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.SqlServer.Management.UI.ConnectionDlg;

namespace SSMSExtension
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// copy vsix to: c:\Program Files\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\Extensions\
    /// while debugging
    /// </remarks>


    //http://www.kendar.org/?p=/tutorials/vsextensions/part04

    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(BackgroundWorksVSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class BackgroundWorksVSPackage : Package
    {
        /// <summary>
        /// BackgroundVSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "9072b500-6b57-4eec-b81a-b1bbfd940421";

        /// <summary>
        /// Initializes a new instance of the <see cref="BackgroundWorksVSPackage"/> class.
        /// </summary>
        public BackgroundWorksVSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        DocumentEvents _documentEvents;


        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            var _dte = (DTE)GetService(typeof(SDTE));
            var _dteEvents = _dte.Events;
            _documentEvents = _dteEvents.DocumentEvents;
            _documentEvents.DocumentOpening += _documentEvents_DocumentOpening;
            _documentEvents.DocumentOpened += _documentEvents_DocumentOpened;


            base.Initialize();
        }

        private void _documentEvents_DocumentOpened(Document document)
        {
            /* not working:
             *  IScriptFactory scriptFactory = ServiceCache.ScriptFactory;
              UIConnectionInfoUtil.SetCustomConnectionColor(scriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo, System.Drawing.Color.Red);
              */

            

        }

        private void _documentEvents_DocumentOpening(string DocumentPath, bool ReadOnly)
        {
            IScriptFactory scriptFactory = ServiceCache.ScriptFactory;
            /*    On Menu > tools > options > debugging > General:
              Ensure "Redirect all output window text to the immediate window" is NOT checked
              */
            Trace.WriteLine(scriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo.ServerName);
            //not working:
            //UIConnectionInfoUtil.SetCustomConnectionColor(scriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo, System.Drawing.Color.Red);

        }

        protected override int QueryClose(out bool canClose)
        {
            UserRegistryRoot.CreateSubKey(@"Packages\{" + PackageGuidString + "}").SetValue("SkipLoading", 1);
            return base.QueryClose(out canClose);
        }


        #endregion
    }
}
