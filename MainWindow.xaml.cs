using Microsoft.Win32;
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
using System.Xml.Linq;

namespace Poster
{
    public class RpoCode
    {
        public string FirmID { get; set; }
        public string BarCode { get; set; }
    }

    public class RpoStatus
    {
        public string BarCode { get; set; }
        public string Status { get; set; }
        public string DateOper { get; set; }
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

        private string _ticketId;

        private void Log(string Text)
        {
            listBox.Items.Add(Text);
        }

        private List<RpoCode> ImportFile(string fileName)
        {
            var Doc = new XmlDocument();
            Doc.Load(fileName);            
            var List = new List<RpoCode>();
            var Codes1 = Doc.SelectNodes("dataroot/RPOF");
            AddCodes(Codes1, List);
            var Codes2 = Doc.SelectNodes("dataroot/RPOFIRMA");
            AddCodes(Codes2, List);
            return List;
        }

        private static void AddCodes(XmlNodeList Codes, List<RpoCode> List)
        {
            for (var k = 0; k < Codes.Count; ++k)
            {
                var Item = Codes.Item(k);
                var Code = new RpoCode();
                Code.FirmID = Item.SelectSingleNode("IDFIRM").InnerText;
                Code.BarCode = Item.SelectSingleNode("BARCODE").InnerText;
                List.Add(Code);
            }
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

            XmlNamespaceManager manager1 = new XmlNamespaceManager(soapEnvelopeXml.NameTable);
            manager1.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            manager1.AddNamespace("pos", "http://fclient.russianpost.org/postserver");
            manager1.AddNamespace("fcl", "http://fclient.russianpost.org");
            var req = soapEnvelopeXml.SelectSingleNode("//soapenv:Envelope//soapenv:Body/pos:ticketRequest/request", manager1);
            req.RemoveAll();

            foreach (var code in Codes)
            {
                var child = soapEnvelopeXml.CreateElement("fcl", "Item", "http://fclient.russianpost.org");
                var barAtt = soapEnvelopeXml.CreateAttribute("Barcode");
                barAtt.Value = code.BarCode;
                child.Attributes.Append(barAtt);
                req.AppendChild(child);
            }

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
                    Log($"Ошибка исполнения запроса: ${ErrorNode.Attributes["ErrorName"].Value}");
                }
                else
                {
                    Log($"Неизвестная ошибка исполнения запроса");
                }
                return null;
            }            
        }

        private List<RpoStatus> ExtractResponse(string ticketId)
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

            var Result = new List<RpoStatus>();
            using (WebResponse response = request.GetResponse())
            {
                var ResponseXml = new XmlDocument();
                ResponseXml.Load(response.GetResponseStream());
                XmlNamespaceManager manager = new XmlNamespaceManager(ResponseXml.NameTable);
                manager.AddNamespace("S", "http://schemas.xmlsoap.org/soap/envelope/");
                manager.AddNamespace("ns2", "http://fclient.russianpost.org/postserver");
                manager.AddNamespace("ns3", "http://fclient.russianpost.org");

                var ErrorNode = ResponseXml.SelectSingleNode("//S:Envelope//S:Body//ns2:ticketResponse/error", manager);
                if (ErrorNode != null)
                {
                    Log($"Ошибка исполнения запроса: ${ErrorNode.Attributes["ErrorName"].Value}");
                }
                var List = ResponseXml.SelectNodes("//S:Envelope//S:Body/ns2:answerByTicketResponse/value//ns3:Item", manager);
                if (List.Count > 0)
                {
                    for (var k = 0; k < List.Count; ++k)
                    {
                        var Item = List.Item(k);
                        var Status = new RpoStatus();
                        Status.BarCode = Item.Attributes["Barcode"].Value;
                        if (Item.LastChild != null)
                        {
                            var OperName = Item.LastChild.Attributes["OperName"];
                            if (OperName != null)
                            {
                                Status.Status = OperName.Value;
                            }
                            else
                            {
                                var ErrorName = Item.LastChild.Attributes["ErrorName"];
                                if (ErrorName != null)
                                {
                                    Status.Status = ErrorName.Value;
                                }
                            }
                            var DateOper = Item.LastChild.Attributes["DateOper"];
                            if (DateOper != null)
                            {
                                Status.DateOper = Item.LastChild.Attributes["DateOper"].Value;
                            }
                            else
                            {
                                Status.DateOper = "Unknown";
                            }
                            Result.Add(Status);
                        }                        
                    }
                }
            }
            return Result;
        }

        private void ProcessFile(string fileName)
        {
            var Codes = ImportFile(fileName);
            Log($"Файл ${fileName} импортирован");
            Log($"Найдено ${Codes.Count} РПО кодов");
            var TicketId = SendRequest(Codes);
            this._ticketId = TicketId;
            if (!string.IsNullOrEmpty(TicketId))
            {
                Log($"Получен код запроса {TicketId}");
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            var Dialog = new Microsoft.Win32.OpenFileDialog();
            Dialog.DefaultExt = ".xml";
            Dialog.Filter = "Xml (*.xml)|*.xml|Text files (*.txt)|*.txt";
            if (Dialog.ShowDialog() == true)
            {
                ProcessFile(Dialog.FileName);
            }            
        }

        private void ExportToXml(List<RpoStatus> items, string fileName)
        {
            var Doc = new XElement("RPO");
            foreach (var Item in items)
            {
                var Elem = new XElement("BARCODE");
                Elem.Add(new XElement("CODE", Item.BarCode));
                Elem.Add(new XElement("STATUS", Item.Status));
                Elem.Add(new XElement("DATE", Item.DateOper));
                Doc.Add(Elem);
            }
            Doc.Save(fileName);
        }

        private void ExportToCsv(List<RpoStatus> items, string fileName)
        {
            //before your loop
            var csv = new StringBuilder();

            csv.AppendLine($"Код;Статус;Дата");
            foreach (var Item in items)
            {
                csv.AppendLine($"_{Item.BarCode};{Item.Status};{Item.DateOper}");
            }

            //after your loop
            File.WriteAllText(fileName, csv.ToString(), Encoding.GetEncoding("windows-1251"));
        }

        private void ExportItems(List<RpoStatus> items)
        {
            SaveFileDialog SaveDialog = new SaveFileDialog();
            SaveDialog.Filter = "Xml|*.xml|CSV|*.csv";
            SaveDialog.Title = "Save an Image File";
            SaveDialog.ShowDialog();
            if (SaveDialog.FileName != "")
            {
                if (System.IO.Path.GetExtension(SaveDialog.FileName).ToLower() == ".csv")
                {
                    ExportToCsv(items, SaveDialog.FileName);
                }
                else
                {
                    ExportToXml(items, SaveDialog.FileName);
                }
            }
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            if (this._ticketId != null) {
                var Items = ExtractResponse(this._ticketId);
                if (Items.Count > 0)
                {
                    ExportItems(Items);
                }
                else
                {
                    System.Windows.MessageBox.Show("Результат обрабатывается. Подождите.");
                }
            }
            else
            {
                Log($"Отсутствует код запроса. Откройте файл для его получения.");
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textLogin.Text = ConfigurationManager.AppSettings["Login"];
            textPass.Text = ConfigurationManager.AppSettings["Password"];
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            AddUpdateAppSettings("Login", textLogin.Text);
            AddUpdateAppSettings("Password", textPass.Text);
        }

        static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }
    }
}
