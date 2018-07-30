using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.ArcMapUI;

namespace ExtensionsReleaser_AddIn
{
    public class ExtensionsReleaser_ArcMap : ESRI.ArcGIS.Desktop.AddIns.Extension
    {
        public ExtensionsReleaser_ArcMap()
        {
        }
        private IApplication m_application;
        protected override void OnStartup()
        {
            try
            {
                m_application = ArcMap.Application;
                if (m_application != null)
                {
                    WireDocumentEvents(ArcMap.Document as ESRI.ArcGIS.Framework.IDocument);
                    //DisableExtensions(m_application);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("There was a problem wiring events: {0}", ex.Message));
                
            }
                     
        }

        protected override void OnShutdown()
        {
            try
            {
                UnwireDocumentEvents();
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("There was a problem unwiring events: {0}", ex.Message));

            }
        }

        #region Document events
        //Event member variable.
        private IDocumentEvents_Event m_docEvents = null;
        private IDocumentEvents_NewDocumentEventHandler m_docNewHandler = null;
        private IDocumentEvents_OpenDocumentEventHandler m_docOpenHandler = null;
        //Wiring.
        private void WireDocumentEvents(ESRI.ArcGIS.Framework.IDocument myDocument)
        {
            m_docEvents = myDocument as IDocumentEvents_Event;
            //Safer wiring
            m_docNewHandler = new IDocumentEvents_NewDocumentEventHandler(OnNewDocument);
            m_docEvents.NewDocument += m_docNewHandler;

            m_docOpenHandler = new IDocumentEvents_OpenDocumentEventHandler(OnOpenDocument);
            m_docEvents.OpenDocument += m_docOpenHandler;
            
        }

        private void UnwireDocumentEvents()
        {
            ArcMap.Events.NewDocument -= m_docNewHandler;
            ArcMap.Events.OpenDocument -= m_docOpenHandler;
        }

        //Event handler methods.
        void OnNewDocument()
        {
            DisableExtensions(m_application);
            UnwireDocumentEvents();
            FindCommandAndExecute(m_application, "esriFramework.ExtensionsCommand");
        }

        void OnOpenDocument()
        {
            DisableExtensions(m_application);
            UnwireDocumentEvents();
            FindCommandAndExecute(m_application, "esriFramework.ExtensionsCommand");

        }
        #endregion
        
        public void DisableExtensions(IApplication app)
        {
            try
            {
                IExtensionManager pExtManager = app as IExtensionManager;
                IExtensionConfig pExtConfig;
                IJITExtensionManager jitExtManager = app as IJITExtensionManager;
                //Custom Extensions
                for (int i = 0; i < pExtManager.ExtensionCount; i++)
                {
                    IExtension ext = pExtManager.Extension[i];
                    pExtConfig = ext as IExtensionConfig;
                    if (pExtConfig != null)
                    {
                        pExtConfig.State = esriExtensionState.esriESDisabled;
                    }
                }
                //OOTB Extensions (3d Analyst, Spatial Analyst, etc)
                for (int i = 0; i < jitExtManager.JITExtensionCount; i++)
                {
                    UID extID = jitExtManager.get_JITExtensionCLSID(i);
                    if (jitExtManager.IsExtensionEnabled(extID))
                    {
                        Console.WriteLine("Extension with clsid {0} is enabled", extID.Value);
                        IExtension ext = app.FindExtensionByCLSID(extID);

                        if (ext != null)
                        {
                            pExtConfig = ext as IExtensionConfig;
                            if (pExtConfig != null)
                                pExtConfig.State = esriExtensionState.esriESDisabled;

                        }
                    }

                }
            }
            catch (Exception ex)
            {
                //Write to log. 
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
