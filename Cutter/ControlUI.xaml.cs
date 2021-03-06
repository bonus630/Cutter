using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using corel = Corel.Interop.VGCore;

namespace Cutter
{
    public partial class ControlUI : UserControl
    {
        private Addon addon;
        private corel.Application corelApp;
        private Styles.StylesController stylesController;
        public ControlUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                
                addon = new Addon(this.corelApp);
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }
            btn_Command.Click += (s, e) =>
            {
                try
                {
                    this.corelApp.EventsEnabled = false;
                    addon.test2();
                }
                catch(Exception erro)
                {
                    this.corelApp.MsgShow(erro.Message);
                }
                finally
                {
                    this.corelApp.EventsEnabled = true;
                }
            };
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
            //this.corelApp.OpenDocument("D:\\temp\\testeZoomTool.cdr");
        }

    }
}
