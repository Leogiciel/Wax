//------------------------------------------------------------------------------
// <copyright file="RememberPropertiesEditor.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace WixRememberPropertiesEditor
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;
    using tomenglertde.Wax.Model.VisualStudio;
    using System.Collections.Generic;
    using tomenglertde.Wax.Model.Wix;
    using System.Linq;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("2f43b37e-07c8-4958-914b-86ff71f8c70c")]
    public class RememberPropertiesEditor : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RememberPropertiesEditor"/> class.
        /// </summary>
        public RememberPropertiesEditor() : base(null)
        {
            this.Caption = "Remember Properties Editor";
            var t = typeof(EnvDTE.Solution);
            if (t.GetMembers().Length == 0)
            {
                // Just to make sure the assembly is loaded, loading it dynamically from XAML may not work!
            }
            var dte = (EnvDTE.DTE)GetService(typeof(EnvDTE.DTE));
            var dte2 = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE)) as DTE;

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new RememberPropertiesEditorControl(dte2);
        }
    }
}
