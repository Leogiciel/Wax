//------------------------------------------------------------------------------
// <copyright file="RememberPropertiesEditorControl.xaml.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace WixRememberPropertiesEditor
{
    using EnvDTE;
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using tomenglertde.Wax.Model.VisualStudio;
    using tomenglertde.Wax.Model.Wix;

    /// <summary>
    /// Interaction logic for RememberPropertiesEditorControl.
    /// </summary>
    public partial class RememberPropertiesEditorControl : UserControl
    {
        #region Private Fields

        private DTE _dte;
        private EnvDTE.Project _project;
        private IEnumerable<EnvDTE.Project> _projects;
        private ProjectItem _propertiesFolder;

        private EnvDTE.Solution _solution;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RememberPropertiesEditorControl"/> class.
        /// </summary>
        public RememberPropertiesEditorControl()
        {
            this.InitializeComponent();
        }

        public RememberPropertiesEditorControl(DTE dte)
        {
            _dte = dte;
            DataContext = this;
            this.InitializeComponent();
            RefreshProjects();
        }

        #endregion Public Constructors

        #region Public Properties

        public ObservableCollection<EnvDTE.Project> Projects
        {
            get
            {
                return new ObservableCollection<EnvDTE.Project>(_projects);
            }
        }

        public List<WixPropertyModel> Properties
        {
            get;
            set;
        }

        public EnvDTE.Project SelectedProject
        {
            get
            {
                return _project;
            }
            set
            {
                _project = value;
                RefreshProperties();
            }
        }

        #endregion Public Properties

        #region Private Properties

        private tomenglertde.Wax.Model.Wix.WixProject _wixProject
        {
            get
            {
                return new tomenglertde.Wax.Model.Wix.WixProject(_wixSolution, _project);
            }
        }

        private tomenglertde.Wax.Model.VisualStudio.Solution _wixSolution
        {
            get
            {
                return new tomenglertde.Wax.Model.VisualStudio.Solution(_solution);
            }
        }

        #endregion Private Properties

        #region Private Methods

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            SetProperties();
        }

        private ProjectItem GetPropertiesProjectItem()
        {
            ProjectItem result = null;
            for (var i = 1; i <= _project.ProjectItems.Count; i++)
            {
                var currentItem = _project.ProjectItems.Item(i);
                if (currentItem != null && currentItem.Kind == "{6bb5f8ef-4483-11d3-8bcf-00c04f8ec28c}" && currentItem.Name == "Properties")
                {
                    result = currentItem;
                    break;
                }
            }
            return result;
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshProjects();
        }

        private void RefreshProjects()
        {
            _solution = _dte.Solution;
            _projects = _solution.GetProjects()
                .Where(project => project.Kind.Equals("{930c7802-8a8c-48f9-8165-68863bccd9dd}", StringComparison.OrdinalIgnoreCase));
            comboBox.ItemsSource = null;
            comboBox.ItemsSource = Projects;
        }

        private void RefreshProperties()
        {
            Properties = new List<WixPropertyModel>();
            if (SelectedProject != null)
            {
                _propertiesFolder = GetPropertiesProjectItem();
                if (_propertiesFolder != null)
                {
                    foreach (ProjectItem file in _propertiesFolder.ProjectItems)
                    {
                        var propertyName = file.Name.Remove(file.Name.LastIndexOf(".wxs"), 4);
                        var matchingProp = _wixProject.ProductNode.PropertyNodes.FirstOrDefault(wp => wp.Id == propertyName.ToUpper())?.Property;
                        if (matchingProp != null)
                            Properties.Add(new WixPropertyModel(propertyName, matchingProp.Value));
                    }
                }
                else MessageBox.Show("Project must contains a \"Properties\" folder");
            }
            dataGrid.ItemsSource = null;
            dataGrid.ItemsSource = Properties;
        }

        private void SetProperties()
        {
            var templateFile = Path.Combine(Path.GetDirectoryName(GetType().Assembly.Location), "Resources", "PropertyName.wxs");
            foreach (var property in Properties.Where(p => !string.IsNullOrWhiteSpace(p.Name)))
            {
                try
                {
                    //Generate wxs file from template
                    var newFileName = templateFile.Replace("PropertyName", property.Name);
                    string text = File.ReadAllText(templateFile);
                    text = text.Replace("PropertyName", property.Name);
                    text = text.Replace("PROPERTYNAME", property.Name.ToUpper());
                    //Add or update wxs file
                    bool needToAdd = true;
                    for (var i = 1; i <= _propertiesFolder.ProjectItems.Count; i++)
                    {
                        ProjectItem existing = _propertiesFolder.ProjectItems.Item(i);
                        if (existing != null && existing.Name == Path.GetFileName(newFileName))
                        {
                            newFileName = existing.Document.FullName;
                            needToAdd = false;
                            break;
                        }
                    }
                    File.WriteAllText(newFileName, text);
                    if (needToAdd)
                        _propertiesFolder.ProjectItems.AddFromFile(newFileName);
                    //Add property in product node
                    var matchingProp = _wixProject.ProductNode.PropertyNodes.FirstOrDefault(wp => wp.Id == property.Name.ToUpper());
                    if (matchingProp != null)
                        matchingProp.Remove();
                    var propertyNode = _wixProject.ProductNode.AddProperty(property.GetProperty());
                    //Add registry search in property
                    propertyNode.AddRegistrySearch(new WixRegistrySearch
                    {
                        Id = $"Remember{property.Name}",
                        Root = RegistrySearchRootType.HKLM,
                        Key = "SOFTWARE\\[Manufacturer]\\[ProductName]",
                        Name = property.Name,
                        SearchType = RegistrySearchType.raw,
                        Win64 = true
                    });
                    //Add custom action ref if needed
                    var customActionRefName = $"Save{property.Name}CmdLineValue";
                    if (!_wixProject.ProductNode.EnumerateCustomActionRefs().Any(car => car == customActionRefName))
                        _wixProject.ProductNode.AddCustomActionRef(customActionRefName);
                    //Add component group ref in feature
                    _wixProject.ForceFeatureRef($"{property.Name}Components");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Exception");
                }
                
            }
            RefreshProperties();
        }

        #endregion Private Methods
    }
}