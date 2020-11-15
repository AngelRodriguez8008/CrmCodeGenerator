using CrmCodeGenerator.VSPackage.Helpers;
using EnvDTE;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace CrmCodeGenerator.VSPackage.Dialogs
{
    /// <summary>
    /// Interaction logic for AddTemplate.xaml
    /// </summary>
    public partial class AddTemplate : Microsoft.VisualStudio.PlatformUI.DialogWindow
    {
        private readonly AddTemplateProp _Props;
        public AddTemplateProp Props => _Props;

        private bool _Canceled = true;
        public bool Canceled => _Canceled;

        public AddTemplate(EnvDTE80.DTE2 dte, Project project)
        {
            WifDetector.CheckForWifInstall();

            InitializeComponent();

            var main = dte.GetMainWindow();
            Owner = main;
            //Loaded += delegate { this.CenterWindow(main); };

            _Props = new AddTemplateProp();
            DataContext = Props;

            var samplesPath = Path.Combine(DteHelper.AssemblyDirectory(), @"Resources\Templates");
            var dir = new DirectoryInfo(samplesPath);
            Props.TemplateList = new ObservableCollection<string>(dir.GetFiles().Select(x => x.Name).Where(x => !x.Equals("Blank.tt")).ToArray());
            Props.Template = "CrmSchema.tt";
            Props.Folder = project.GetProjectDirectory();
            
        }
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            this.HideMinimizeAndMaximizeButtons();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DefaultTemplate.SelectedIndex = 0;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            //this.DialogResult = false;
            _Canceled = true;
            Close();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            //this.DialogResult = true;
            _Canceled = false;
            Close();
        }
    }

    public class AddTemplateProp : INotifyPropertyChanged
    {
        // boiler-plate
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            handler?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        //protected bool SetField<T>(ref T field, T value, string propertyName)
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }


        private string _Template;
        public string Template
        {
            get => _Template;
            set
            {
                SetField(ref _Template, value);
                NewTemplate = !File.Exists(Path.Combine(_Folder, _Template));
            }
        }
        private string _Folder = "";
        public string Folder
        {
            get => _Folder;
            set => SetField(ref _Folder, value);
        }

        private bool _NewTemplate;
        public bool NewTemplate
        {
            get => _NewTemplate;
            set => SetField(ref _NewTemplate, value);
        }

        private string _OutputPath;
        public string OutputPath
        {
            get => _OutputPath;
            set => SetField(ref _OutputPath, value);
        }

        private ObservableCollection<string> _TemplateList = new ObservableCollection<string>();
        public ObservableCollection<string> TemplateList
        {
            get => _TemplateList;
            set => SetField(ref _TemplateList, value);
        }
    }
}
