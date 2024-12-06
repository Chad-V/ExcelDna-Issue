//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;

namespace MyAddIn
{
	public interface IPaneHolder {
		public ElementHost Host { get;}

		public void SetRootComponent(System.Windows.UIElement root);
	}

	//[ComVisible(true)]
	[ComDefaultInterface(typeof(IPaneHolder))]
	public partial class PaneHolder : System.Windows.Forms.UserControl, IPaneHolder
	{
		public ElementHost Host { get; private set; }

		public PaneHolder()
		{
			InitializeComponent();

			// Construct the ElementHost component, used to host any WPF
			// views within the WinForms component. Needed to display
			// WPF within the Excel popups, which are WinForms.
			Host = new();
			Host.Dock = DockStyle.Fill;
			Controls.Add(Host);
		}

		public void SetRootComponent(System.Windows.UIElement root)
		{
			// Set the WPF root component.
			Host.Child = root;
		}
	}
}
