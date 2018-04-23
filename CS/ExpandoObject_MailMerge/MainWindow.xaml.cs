using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Windows;
using System.Xml.Linq;

namespace ExpandoObject_MailMerge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);

        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            richEditControl1.ApplyTemplate();

            dynamic weathers = GetExpandoFromXml("weather.xml");
            richEditControl1.Options.MailMerge.DataSource = weathers;

            richEditControl1.LoadDocument("weather_report.rtf");

            ribbonControl1.SelectedPage = pageMailings;
            richEditControl1.Options.MailMerge.ViewMergedData = true;
        }

        public static IList<dynamic> GetExpandoFromXml(String file)
        {
            var weathers = new List<dynamic>();

            var doc = XDocument.Load(file);
            var nodes = from node in doc.Root.Descendants("weather")
                        select node;
            foreach (var n in nodes) {
                dynamic MyData = new ExpandoObject();
                MyData.LastUpdateTime = String.Format("{0:o}", DateTime.Now);
                MyData.Weather = new ExpandoObject();
                foreach (var child in n.Descendants()) {

                    var w = MyData.Weather as IDictionary<String, object>;
                    XAttribute atb = child.Attribute("data");
                    if (atb != null)
                        w[child.Name.LocalName] = atb.Value;
                }

                weathers.Add(MyData);

            }
            return weathers;
        }

    }
}
