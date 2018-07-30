using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;

namespace ExtensionsReleaser_AddIn
{
    public class ExtensionsReleaser_ArcCatalog : ESRI.ArcGIS.Desktop.AddIns.Extension
    {
        public ExtensionsReleaser_ArcCatalog()
        {
        }
        private IApplication m_application;
        protected override void OnStartup()
        {

            m_application = ArcCatalog.Application;
            if (m_application != null)
                SetUpArcCatalogEvents(ArcCatalog.Document as ESRI.ArcGIS.Framework.IDocument);
        }

        private void SetUpArcCatalogEvents(IDocument iDocument)
        {
            Debug.WriteLine("Setting up events");
            ArcCatalog.Events.NewDocument += delegate() { ArcCatalog_NewDocument(); };
            ArcCatalog.Events.OpenDocument += delegate() { ArcCatalog_OpenDocument(); };
        }

        private void ArcCatalog_OpenDocument()
        {
            DisableExtensions(m_application);
            UnwireDocumentEvents();
            //FindCommandAndExecute(m_application, "esriFramework.ExtensionsCommand");

        }

        private void ArcCatalog_NewDocument()
        {
            DisableExtensions(m_application);
            UnwireDocumentEvents();
            //FindCommandAndExecute(m_application, "esriFramework.ExtensionsCommand");

        }

        private void UnwireDocumentEvents()
        {
            ArcCatalog.Events.NewDocument -= ArcCatalog_NewDocument;
            ArcCatalog.Events.OpenDocument -= ArcCatalog_OpenDocument;
        }
        public void DisableExtensions(IApplication app)
        {
            try
            {
                IExtensionManager pExtManager = app as IExtensionManager;
                IExtensionConfig pExtConfig;
                IJITExtensionManager jitExtManager = app as IJITExtensionManager;


                for (int i = 0; i < pExtManager.ExtensionCount; i++)
                {
                    IExtension ext = pExtManager.Extension[i];
                    pExtConfig = ext as IExtensionConfig;
                    if (pExtConfig != null)
                    {
                        pExtConfig.State = esriExtensionState.esriESDisabled;
                    }
                }
                for (int i = 0; i < jitExtManager.JITExtensionCount; i++)
                {
                    UID extID = jitExtManager.get_JITExtensionCLSID(i);
                    IExtension ext = app.FindExtensionByCLSID(extID);

                    if (ext != null)
                    {
                        pExtConfig = ext as IExtensionConfig;
                        if (pExtConfig != null)
                            pExtConfig.State = esriExtensionState.esriESDisabled;

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Extensions could not be released: {0}", ex.Message));
            }
        }
        ///<summary>Find a command and click it programmatically.</summary>
        ///  
        ///<param name="application">An IApplication interface.</param>
        ///<param name="commandName">A System.String that is the name of the command to return. Example: "esriFramework.HelpContentsCommand" or "{D74B2F25-AC90-11D2-87F8-0000F8751720}"</param>
        ///   
        ///<remarks>Refer to the EDN document http://edndoc.esri.com/arcobjects/9.1/default.asp?URL=/arcobjects/9.1/ArcGISDevHelp/TechnicalDocuments/Guids/ArcMapIds.htm for a listing of available CLSID's and ProgID's that can be used as the commandName parameter.</remarks>
        public void FindCommandAndExecute(ESRI.ArcGIS.Framework.IApplication application, System.String commandName)
        {
            try
            {
                ESRI.ArcGIS.Framework.ICommandBars commandBars = application.Document.CommandBars;
                ESRI.ArcGIS.esriSystem.UID uid = new ESRI.ArcGIS.esriSystem.UIDClass();
                uid.Value = commandName; 
                ESRI.ArcGIS.Framework.ICommandItem commandItem = commandBars.Find(uid, false, false);

                if (commandItem != null)
                {
                    commandItem.Execute();
                }
            }
            catch (Exception ex)
            {
                
                Console.WriteLine(String.Format("Extensions manager could not be opened: {0}", ex.Message));
            }
        }

    }

}
