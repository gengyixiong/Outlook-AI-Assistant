using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Diagnostics;
using OutlookAiAssistant.OutlookIntegration;
using OutlookAiAssistant.Search;
using OutlookAiAssistant.Summary;
using OutlookAiAssistant.UI;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.AddIn
{
    /// <summary>
    /// COM entry point loaded by classic Outlook.
    /// </summary>
    [ComVisible(true)]
    [Guid("8A435B2A-4E46-4557-9300-BD94F8C92F82")]
    [ProgId("OutlookAiAssistant.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class Connect :
        IDTExtensibility2,
        Office.ICustomTaskPaneConsumer,
        Office.IRibbonExtensibility
    {
        private Outlook.Application _application;
        private Office.ICTPFactory _taskPaneFactory;
        private Office.CustomTaskPane _taskPane;
        private Office.IRibbonUI _ribbon;

        public static Connect Current { get; private set; }

        public SettingsStore SettingsStore { get; private set; }
        public AppSettings Settings { get; private set; }
        public OutlookContextService OutlookContext { get; private set; }
        public OutlookSearchService OutlookSearch { get; private set; }
        public SummaryService SummaryService { get; private set; }
        public SearchPlannerService SearchPlanner { get; private set; }
        public AqsQueryCompiler QueryCompiler { get; private set; }
        public SearchPlanFormatter SearchPlanFormatter { get; private set; }

        public void OnConnection(
            object application,
            ext_ConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            try
            {
                _application = application as Outlook.Application;
                if (_application == null)
                {
                    throw new InvalidOperationException(
                        "加载项没有获得 Outlook Application 对象。");
                }

                Current = this;
                SettingsStore = new SettingsStore();
                Settings = SettingsStore.Load();
                OpenAiCompatibleClient aiClient = new OpenAiCompatibleClient();
                SearchPlanParser parser = new SearchPlanParser();
                OutlookContext = new OutlookContextService(_application);
                OutlookSearch = new OutlookSearchService(_application);
                SummaryService = new SummaryService(aiClient, SettingsStore);
                SearchPlanner = new SearchPlannerService(
                    aiClient,
                    SettingsStore,
                    parser);
                QueryCompiler = new AqsQueryCompiler();
                SearchPlanFormatter = new SearchPlanFormatter();
                Logger.Info("Outlook add-in connected.");
            }
            catch (Exception ex)
            {
                Logger.Error("Outlook add-in connection failed.", ex);
                throw;
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            SetTaskPaneVisible(true);
            Logger.Info("Outlook startup completed.");
        }

        public void OnDisconnection(
            ext_DisconnectMode removeMode,
            ref Array custom)
        {
            Logger.Info("Outlook add-in disconnected.");
            _ribbon = null;
            _taskPane = null;
            _taskPaneFactory = null;
            _application = null;
            if (ReferenceEquals(Current, this))
            {
                Current = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Info("Outlook is shutting down.");
        }

        /// <summary>
        /// Office calls this method after its custom task pane factory is ready.
        /// </summary>
        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _taskPaneFactory = CTPFactoryInst;
            EnsureTaskPane();
        }

        public string GetCustomUI(string ribbonID)
        {
            if (string.Equals(
                ribbonID,
                "Microsoft.Outlook.Explorer",
                StringComparison.OrdinalIgnoreCase))
            {
                return BuildRibbonXml("TabMail", "OutlookAiAssistantExplorerGroup");
            }

            if (string.Equals(
                ribbonID,
                "Microsoft.Outlook.Mail.Read",
                StringComparison.OrdinalIgnoreCase))
            {
                return BuildRibbonXml(
                    "TabReadMessage",
                    "OutlookAiAssistantReadGroup");
            }

            return null;
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void OnToggleTaskPane(
            Office.IRibbonControl control,
            bool pressed)
        {
            SetTaskPaneVisible(pressed);
        }

        public bool GetTaskPaneVisible(Office.IRibbonControl control)
        {
            try
            {
                return _taskPane != null && _taskPane.Visible;
            }
            catch
            {
                return false;
            }
        }

        public void SetTaskPaneVisible(bool visible)
        {
            try
            {
                EnsureTaskPane();
                if (_taskPane != null)
                {
                    _taskPane.Visible = visible;
                }

                if (_ribbon != null)
                {
                    _ribbon.Invalidate();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Unable to change task pane visibility.", ex);
            }
        }

        public bool ShowSettings(IWin32Window owner)
        {
            using (SettingsForm form = new SettingsForm(SettingsStore, Settings))
            {
                DialogResult result = owner == null
                    ? form.ShowDialog()
                    : form.ShowDialog(owner);
                if (result != DialogResult.OK)
                {
                    return false;
                }

                Settings = form.SavedSettings;
                return true;
            }
        }

        private void EnsureTaskPane()
        {
            if (_taskPane != null || _taskPaneFactory == null)
            {
                return;
            }

            try
            {
                _taskPane = _taskPaneFactory.CreateCTP(
                    "OutlookAiAssistant.TaskPane",
                    "AI 邮件助手",
                    Type.Missing);
                _taskPane.DockPosition =
                    Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 410;
                _taskPane.Visible = true;
                Logger.Info("Custom task pane created.");
            }
            catch (Exception ex)
            {
                Logger.Error("Custom task pane creation failed.", ex);
            }
        }

        private static string BuildRibbonXml(string tabIdMso, string groupId)
        {
            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\""
                + " onLoad=\"OnRibbonLoad\">"
                + "<ribbon><tabs><tab idMso=\"" + tabIdMso + "\">"
                + "<group id=\"" + groupId + "\" label=\"AI 邮件助手\">"
                + "<toggleButton id=\"" + groupId + "Toggle\""
                + " label=\"显示助手\" size=\"large\" imageMso=\"ResearchPane\""
                + " getPressed=\"GetTaskPaneVisible\""
                + " onAction=\"OnToggleTaskPane\""
                + " screentip=\"显示或隐藏 AI 邮件助手右侧栏\"/>"
                + "</group></tab></tabs></ribbon></customUI>";
        }
    }
}

