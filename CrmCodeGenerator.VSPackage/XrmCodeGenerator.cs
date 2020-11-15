using System;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using CrmCodeGenerator.VSPackage.Model;
using CrmCodeGenerator.VSPackage.T4;
using CrmCodeGenerator.VSPackage.Dialogs;
using Microsoft.VisualStudio.TextTemplating.VSHost;
using Microsoft.VisualStudio.TextTemplating;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.VSPackage
{
    // http://blogs.msdn.com/b/vsx/archive/2013/11/27/building-a-vsix-deployable-single-file-generator.aspx
    [Guid(GuidList.guidCrmCodeGenerator_SimpleGenerator)]
    public class XrmCodeGenerator : BaseCodeGenerator
    {
        public const string Name = nameof(XrmCodeGenerator);
        public const string Description = "Dynamics CRM Code Generator for Visual Studio";

        private readonly Settings settings = Configuration.Instance.Settings;
        private string extension;
        private Mapper mapper;

        public override string GetDefaultExtension() => extension ?? "cs";

        public string GenerateCode(string inputFileName)
        {
            var resultCode = BuildCode(inputFileName, File.ReadAllText(inputFileName));
            return resultCode;
        }

        protected override byte[] GenerateCode(string inputFileName, string inputFileContent)
            => Encoding.UTF8.GetBytes(BuildCode(inputFileName, inputFileContent));
        
        public string BuildCode(string inputFileName, string inputFileContent)
        {
            try
            {
                if (inputFileContent == null)
                    throw new ArgumentException(inputFileContent);

                Status.Update(Environment.NewLine  + "CRM Code Generator -->> START");
                settings.IsActive = true;

                Status.Update("Loading Template File... ");

                Status.Update("Loading Entities & Attributes Mapping File... ");

                var mappingFile = Path.ChangeExtension(inputFileName, "mapping.json");
                Status.Update($"Mapping File: {mappingFile ?? "null"}");
                if (string.IsNullOrWhiteSpace(mappingFile) == false)
                {
                    settings.MappingSettings = LoadMappingFromFile(mappingFile);

                    var entities = settings.MappingSettings?.Entities?.Keys.ToArray();
                    if (entities != null)
                    {
                        var logicalNames = string.Join(",", entities);

                        Status.Update($"Entities Mapping: {logicalNames}");
                        settings.EntitiesToIncludeString = logicalNames;
                    }
                }

                PromptToRefreshEntities();

                if (mapper == null)
                {
                    var dte = Package.GetGlobalService(typeof(SDTE)) as EnvDTE80.DTE2;
                    var m = new Login(dte, settings);
                    m.ShowDialog();
                    mapper = m.Mapper;

                    if (mapper == null)
                    {
                        Status.Update("Cancelled by the user");
                        return null;
                    }
                }

                Status.Update("Creating Mapping Context... ");
                Context context = mapper.CreateContext();
                Status.Update("Generating code from template... ");
                var t4 = Package.GetGlobalService(typeof(STextTemplating)) as ITextTemplating;
                var sessionHost = t4 as ITextTemplatingSessionHost;
                if (sessionHost == null)
                {
                    var error = "Unexpected Error occur by Initializing the SessionHost. Abort";
                    Status.Update(error);
                    return error;
                }

                context.Namespace = FileNamespace;
                sessionHost.Session = sessionHost.CreateSession();
                sessionHost.Session["Context"] = context;

                var cb = new Callback();
                t4.BeginErrorSession();
                string content = t4.ProcessTemplate(inputFileName, inputFileContent, cb);
                t4.EndErrorSession();

                // If there was an output directive in the TemplateFile, then cb.SetFileExtension() will have been called.
                if (!string.IsNullOrWhiteSpace(cb.FileExtension))
                {
                    extension = cb.FileExtension;
                }
                
                if (cb.ErrorMessages.Count > 0)
                {
                    var errors = new StringBuilder();
                    //Configuration.Instance.DTE.ExecuteCommand("View.ErrorList");
                    // Append any error/warning to output window
                    foreach (var err in cb.ErrorMessages)
                    {
                        // The templating system (eg t4.ProcessTemplate) will automatically add error/warning to the ErrorList 
                        var errorLine = "[" + (err.Warning ? "WARN" : "ERROR") + "] " + err.Message + " " + err.Line + "," + err.Column;
                        errors.AppendLine(errorLine);
                        Status.Update(errorLine);
                    }
                    var error = errors.ToString();
                    Status.Update(error);
                    return error;
                }

                Status.Update("Writing code to disk... ");
                Status.Update("Done!");
                return content;
            }
            catch (Exception ex)
            {
                var error = "[ERROR] " + ex.Message + (ex.InnerException != null ? "\n" + "[ERROR] " + ex.InnerException.Message : "");
                Status.Update(error);
                Status.Update(ex.StackTrace);
                Status.Update("Unable to map entities, see error above.");
                return error + Environment.NewLine + ex.StackTrace;
            }
        }
        
        public static MappingSettings LoadMappingFromFile(string mappingFile)
        {
            Status.Update("Loading Mapping from File ...");
            bool exist = File.Exists(mappingFile);
            if (exist == false)
                return null;

            var json = File.ReadAllText(mappingFile);
            if (string.IsNullOrWhiteSpace(json))
                return null;

            var mapping = Deserialize<MappingSettings>(json);
            return mapping;
        }

        public static T Deserialize<T>(string json) where T : class
        {
            if (string.IsNullOrEmpty(json))
                return default(T);

            var settings = new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true };
            using (var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                var serializer = new DataContractJsonSerializer(typeof(T), settings);
                return serializer.ReadObject(memoryStream) as T;
            }
        }

        private void PromptToRefreshEntities()
        {
            if (mapper == null)
                return;

            ThreadHelper.ThrowIfNotOnUIThread();

            var results = VsShellUtilities.ShowMessageBox(ServiceProvider.GlobalProvider, "Do you want to refresh the CRM Entities from the Server?", "Refresh", OLEMSGICON.OLEMSGICON_QUERY, OLEMSGBUTTON.OLEMSGBUTTON_YESNO, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            if (results == 6)
                mapper = null;
        }
    }
}
