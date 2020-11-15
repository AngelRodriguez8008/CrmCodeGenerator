using CrmCodeGenerator.VSPackage.Model;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextTemplating;
using Microsoft.VisualStudio.TextTemplating.VSHost;
using System.Linq;

namespace CrmCodeGenerator.VSPackage.T4
{
    class Processor
    {

        public static string ProcessTemplateCore(string templatePath, string templateContent, Context context, out string extension)
        {
            extension = null;

            // Get the text template service:
            var t4 = Package.GetGlobalService(typeof(STextTemplating)) as ITextTemplating;
            var sessionHost = t4 as ITextTemplatingSessionHost;
            if (sessionHost == null)
                return null;

            // Create a Session in which to pass parameters:
            sessionHost.Session = sessionHost.CreateSession();
            sessionHost.Session["Context"] = context;

            // string templateContent = System.IO.File.ReadAllText(templatePath);
            var cb = new Callback();

            // Process a text template:
            string result = t4.ProcessTemplate(templatePath, templateContent, cb);

            // If there was an output directive in the TemplateFile, then cb.SetFileExtension() will have been called.
            extension = string.IsNullOrWhiteSpace(cb.FileExtension) 
                ? ".cs" 
                : cb.FileExtension;

            // Append any error messages:
            if (cb.ErrorMessages.Count > 0)
            {
                result = cb.ErrorMessages.ToString();
            }
            return result;
        }

        public static string ProcessTemplate(string templatePath, Context context)
        {
            // Get the text template service:
            var t4 = Package.GetGlobalService(typeof(STextTemplating)) as ITextTemplating;
            var sessionHost = t4 as ITextTemplatingSessionHost;
            if (sessionHost == null)
                return null;

            // Create a Session in which to pass parameters:
            sessionHost.Session = sessionHost.CreateSession();
            sessionHost.Session["Context"] = context;

            string templateContent = System.IO.File.ReadAllText(templatePath);
            var cb = new Callback();

            // Process a text template:
            string result = t4.ProcessTemplate(templatePath, templateContent, cb);
            
            var extension = string.IsNullOrWhiteSpace(cb.FileExtension) 
                ? ".cs" 
                : cb.FileExtension;
            var outputFullPath = System.IO.Path.ChangeExtension(templatePath, extension);


            // Write the processed output to file:
            // UpdateStatus("Writing......", true);
            System.IO.File.WriteAllText(outputFullPath, result, cb.OutputEncoding);

            // Append any error messages:
            if (cb.ErrorMessages.Count > 0)
            {
                System.IO.File.AppendAllLines(outputFullPath, cb.ErrorMessages.Select(x => x.Message));
            }

            string errroMessage = null;
            if (cb.ErrorMessages.Count > 0)
            {
                errroMessage = "Unable to generate file see " + outputFullPath + " for more details ";
            }
            return errroMessage;
        }
    }
}
