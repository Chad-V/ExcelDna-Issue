using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using UserControl = System.Windows.Controls.UserControl;
using MessageBox = System.Windows.Forms.MessageBox;

namespace MyAddIn
{
    //public interface IPane<TRoot> : CustomTaskPane
    //	where TRoot : System.Windows.Controls.UserControl
    //{
    //}

    public class CustomPane
    {
        public PaneHolder PaneHolder { get; set; }
        public UserControl RootComponent { get; set; }
        public CustomTaskPane TaskPane { get; set; }

        //public Action<CustomTaskPane> VisibilityCallback { get; set; }
        //public Action<CustomTaskPane> DockPositionChangeCallback { get; set; }

        public string PaneTitle { get; set; }

        public event EventHandler? VisibilityChanged;

        public event EventHandler? DockPositionChanged;

        public CustomPane(UserControl root, string title)
        {
            RootComponent = root;
            PaneTitle = title;

            PaneHolder = new();
            PaneHolder.SetRootComponent(RootComponent);

            TaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(PaneHolder, PaneTitle);
            TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
            TaskPane.Visible = true;

            //TaskPane.

            //VisibilityCallback = DefaultVisibilityCallback;
            //DockPositionChangeCallback = DefaultDockPositionCallback;
        }

        public void SetDockPosition(MsoCTPDockPosition pos)
        {
            TaskPane.DockPosition = pos;
        }

        public void Show()
        {
            TaskPane.Visible = true;
        }

        public void Hide()
        {
            TaskPane.Visible = false;
        }
    }

    public static class CustomTaskPaneManager
    {
        // Static list of all the currently-built panes.
        static List<CustomPane> Panes = new();

        //private static ResourceDictionary dict = new ResourceDictionary() {
        //	Source = new Uri("Content/AppResources.xaml", UriKind.Relative)
        //};

        private static ResourceDictionary? dict;

        private static ResourceDictionary Dictionary
        {
            get
            {
                if (dict is null)
                {
                    Application.ResourceAssembly = Assembly.GetExecutingAssembly();

                    dict = new ResourceDictionary()
                    {
                        Source = new Uri("Content/Default.xaml", UriKind.Relative)
                    };
                }

                return dict;
            }
        }

        public static CustomPane CreateCustomTaskPane(UserControl rootComponent, string title)
        {
            InjectResourcesIntoComponent(rootComponent);

            CustomPane pane = new(rootComponent, title);
            Panes.Add(pane);

            //pane.Show();

            return pane;
        }

        // Helper method to automatically inject required XAML resources into the
        // newly created XAML root components.
        private static void InjectResourcesIntoComponent(UserControl component)
        {
            component.Resources.MergedDictionaries.Add(Dictionary);
        }

        //public static void HidePane()

        // TODO How to hold all of the internal CTPs?
        // Each of them will be a unique custom task pane that is childed to a PaneHolder.
        // They'll all have a unique type for their RootComponent.
        // There should be vis/invis support, as well as deletion.
        // Should there be a new class that stores the root component, the pane holder, and the CTP? And
        // perchance a name or ID of this pane?
    }
}
