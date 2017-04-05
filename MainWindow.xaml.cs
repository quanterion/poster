using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
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
using System.Xml;

namespace Poster
{
    public class RpoCode
    {
        public string FirmID { get; set; }
        public string BarCode { get; set; }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Log(string Text)
        {
            logBlock.Inlines.Add(new Line {
                X1 = 0,
                Y1 = 0,
                X2 = 100,
                Y2 = 0,
                Stroke = new SolidColorBrush(Colors.Black), StrokeThickness = 4.0 }
            );
        }

        private List<RpoCode> ImportFile(string fileName)
        {
            var Doc = new XmlDocument();
            Doc.Load(fileName);
            var Codes = Doc.SelectNodes("dataroot/RPOF");
            var List = new List<RpoCode>();
            for (var k = 0; k < Codes.Count; ++k)
            {
                var Item = Codes.Item(k);
                var Code = new RpoCode();
                Code.FirmID = Item.SelectSingleNode("IDFIRM").InnerText;
                Code.BarCode = Item.SelectSingleNode("BARCODE").InnerText;
                List.Add(Code);
            }                
            return List;
        }

        private string SendRequest(List<RpoCode> Codes)
        {
            var Url = "https://tracking.russianpost.ru/fc";
            var request = (HttpWebRequest)WebRequest.Create(Url);
            request.Headers.Add(@"SOAP:Action");
            request.ContentType = "text/xml;charset=\"utf-8\"";
            request.Accept = "text/xml";
            request.Method = "POST";

            var Login = ConfigurationManager.AppSettings["Login"];
            var Password = ConfigurationManager.AppSettings["Password"];

            XmlDocument soapEnvelopeXml = new XmlDocument();
            soapEnvelopeXml.LoadXml($@"
                <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/""
                                    xmlns:pos=""http://fclient.russianpost.org/postserver""
                                    xmlns:fcl=""http://fclient.russianpost.org"">
                    <soapenv:Header/>
                    <soapenv:Body>
                        <pos:ticketRequest>
                            <request>
                                <fcl:Item Barcode=""RA123456788RU""/>
                                <fcl:Item Barcode=""RA123456789RU""/>
                                <fcl:Item Barcode=""RA123456780RU""/>
                            </request>
                            <login>{Login}</login>
                            <password>{Password}</password>
                            <language>RUS</language>
                        </pos:ticketRequest>
                    </soapenv:Body>
                </soapenv:Envelope>");
            using (var stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }
            using (WebResponse response = request.GetResponse())
            {
                var ResponseXml = new XmlDocument();                
                ResponseXml.Load(response.GetResponseStream());
                XmlNamespaceManager manager = new XmlNamespaceManager(ResponseXml.NameTable);
                manager.AddNamespace("S", "http://schemas.xmlsoap.org/soap/envelope/");
                manager.AddNamespace("ns2", "http://fclient.russianpost.org/postserver");
                var Node = ResponseXml.SelectSingleNode("//S:Envelope//S:Body/ns2:ticketResponse/value", manager);
                if (Node != null)
                {
                    return Node.InnerText;
                }
                var ErrorNode = ResponseXml.SelectSingleNode("//S:Envelope//S:Body//ns2:ticketResponse/error", manager);
                if (ErrorNode != null)
                {
                    Log($"Ошибка исполнения запроса: ${ErrorNode.Attributes["ErrorName"]}");
                }
                else
                {
                    Log($"Неизвестная ошибка исполнения запроса");
                }
                return null;
            }            
        }

        private void ExtractResponse(string ticketId)
        {
            var Url = "https://tracking.russianpost.ru/fc";
            var request = (HttpWebRequest)WebRequest.Create(Url);
            request.Headers.Add(@"SOAP:Action");
            request.ContentType = "text/xml;charset=\"utf-8\"";
            request.Accept = "text/xml";
            request.Method = "POST";

            var Login = ConfigurationManager.AppSettings["Login"];
            var Password = ConfigurationManager.AppSettings["Password"];

            XmlDocument soapEnvelopeXml = new XmlDocument();
            soapEnvelopeXml.LoadXml($@"
                <soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/""
                            xmlns:pos=""http://fclient.russianpost.org/postserver"">
                    <soapenv:Header/>
                    <soapenv:Body>
                        <pos:answerByTicketRequest>
                            <ticket>{ticketId}</ticket>
                            <login>{Login}</login>
                            <password>{Password}</password>
                        </pos:answerByTicketRequest>
                    </soapenv:Body>
                </soapenv:Envelope>");
            using (var stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }
            using (WebResponse response = request.GetResponse())
            {
                var ResponseXml = new XmlDocument();
                ResponseXml.Load(response.GetResponseStream());
                XmlNamespaceManager manager = new XmlNamespaceManager(ResponseXml.NameTable);
                manager.AddNamespace("S", "http://schemas.xmlsoap.org/soap/envelope/");
                manager.AddNamespace("ns2", "http://fclient.russianpost.org/postserver");

                // TODO parse response
                var Node = ResponseXml.SelectSingleNode("//S:Envelope//S:Body/ns2:ticketResponse/value", manager);
                if (Node != null)
                {
                    
                }
            }
        }

        private void ProcessFile(string fileName)
        {
            var Codes = ImportFile(fileName);
            Log($"Файл ${fileName} импортирован");
            Log($"Найдено ${Codes.Count} РПО кодов");
            var TicketId = SendRequest(Codes);
            if (TicketId != null)
            {

            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            SendRequest(null);
            var Dialog = new Microsoft.Win32.OpenFileDialog();
            Dialog.DefaultExt = ".xml";
            Dialog.Filter = "Xml (*.xml)|*.xml|Text files (*.txt)|*.txt";
            if (Dialog.ShowDialog() == true)
            {
                ProcessFile(Dialog.FileName);
            }            
        }
    }
}
