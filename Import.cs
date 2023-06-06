using Soneta.Business;
using Soneta.Business.App;
using Soneta.Core;
using Soneta.Towary;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace ZapromImportTowaryXML
{
    internal class Import
    {
        public const string USER = "Administrator";
        public const string PASSWORD = "ol32uiop";
        static Login login;

        public static void SonetaLoader()
        {
            Soneta.Start.Loader loader = new Soneta.Start.Loader { WithExtensions = true };
            loader.Load();
        }

        static void Main(string[] args)
        {
            SonetaLoader();
            ImportujTowary();
        }    

        public static void ImportujTowary()
        {           
            string fromPath = @"D:\ERP\";
            string toPath = @"D:\ERP\ARCHIWUM\";
            Database database = BusApplication.Instance["ZAPROM"];

            /*string fromPath = @"C:\Users\User\Downloads\Zaprom\";
            string toPath = @"C:\Users\User\Downloads\Zaprom\ARCHIWUM\";
            Database database = BusApplication.Instance["DEMO BARTEK"];*/

            login = database.Login(false, USER, PASSWORD);

            DirectoryInfo dir = new DirectoryInfo(fromPath);
            FileInfo[] xmlFiles = dir.GetFiles("*.xml");

            // read rewizje
            IDictionary<string, string> rewizje = new Dictionary<string, string>();

            foreach (FileInfo file in xmlFiles)
            {
                string filePath = file.FullName;

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(filePath);

                XmlNodeList revision = xmlDoc.GetElementsByTagName("Z5_ZapromDesignRevision");

                foreach (XmlNode node in revision)
                {
                    string id = node.Attributes["item_revision_id"].Value;
                    string obj = node.Attributes["object_name"].Value;

                    if(rewizje.ContainsKey(id))
                        rewizje.Remove(id);

                    rewizje.Add(id, obj);
                }
            }

            using (Session session = login.CreateSession(false, false))
            {
                TowaryModule towaryModule = TowaryModule.GetInstance(session);
                CoreModule coreModule = CoreModule.GetInstance(session);
                Towar temp, towar;
                DefinicjaStawkiVat defStawkiVat = coreModule.DefStawekVat.WgKodu["-"];

                foreach (FileInfo file in xmlFiles)
                {                  
                    string filePath = file.FullName;

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(filePath);

                    XmlNodeList zapromDesign = xmlDoc.GetElementsByTagName("Z5_ZapromDesign");
                    XmlNodeList zapromStd = xmlDoc.GetElementsByTagName("Z5_ZapromStd");
                    XmlNodeList units = xmlDoc.GetElementsByTagName("UnitOfMeasure");

                    // read jednostki
                    IDictionary<string, string> jednostki = new Dictionary<string, string>();

                    foreach (XmlNode node in units)
                    {
                        string id = node.Attributes["elemId"].Value;
                        string symbol = node.Attributes["symbol"].Value;
                        jednostki.Add(id, symbol);
                    }

                    // merged list
                    var nodeList = zapromDesign.Cast<XmlNode>().Concat(zapromStd.Cast<XmlNode>());                 

                    foreach (XmlNode node in nodeList)
                    {
                        string kod = node.Attributes["item_id"].Value;
                        string nazwa = node.Attributes["object_name"].Value;                     
                        string idJednostki = node.Attributes["uom_tag"].Value.Replace("#","");
                        string kodJednostki = null;
                        string nrRewizji = null;

                        // znajdz jednostke po id
                        foreach (KeyValuePair<string, string> elem in jednostki)
                        {
                            if (elem.Key == idJednostki)
                            {
                                kodJednostki = elem.Value;
                                break;
                            }                           
                        }

                        Jednostka j = kodJednostki != null ? towaryModule.Jednostki.WgKodu[kodJednostki] : null;

                        // znajdz rewizje po nazwie
                        foreach (KeyValuePair<string, string> elem in rewizje)
                        {
                            if (elem.Value == nazwa)
                            {
                                nrRewizji = elem.Key;

                                kod = node.Attributes["item_id"].Value + "/" + nrRewizji;
                                temp = towaryModule.Towary.WgKodu[kod];

                                if (temp == null)
                                {
                                    // dodaj towar
                                    using (ITransaction trans = session.Logout(true))
                                    {
                                        towar = new Towar();
                                        towaryModule.Towary.AddRow(towar);

                                        towar.Kod = kod;
                                        towar.Nazwa = nazwa;
                                        towar.Jednostka = j ?? towaryModule.Jednostki.WgKodu["szt"];
                                        towar.DefinicjaStawki = defStawkiVat;

                                        if (kod.Contains("TC-P") || kod.Contains("TC-S") || kod.Contains("TC-A"))
                                            towar.Typ = TypTowaru.Produkt;

                                        trans.Commit();
                                    }
                                }
                            }
                        }

                        nrRewizji = nrRewizji ?? "01";

                        kod = node.Attributes["item_id"].Value + "/" + nrRewizji;
                        temp = towaryModule.Towary.WgKodu[kod];

                        if (temp == null)
                        {
                            // dodaj towar
                            using (ITransaction trans = session.Logout(true))
                            {
                                towar = new Towar();
                                towaryModule.Towary.AddRow(towar);

                                towar.Kod = kod;
                                towar.Nazwa = nazwa;
                                towar.Jednostka = j ?? towaryModule.Jednostki.WgKodu["szt"];
                                towar.DefinicjaStawki = defStawkiVat;

                                if (kod.Contains("TC-P") || kod.Contains("TC-S") || kod.Contains("TC-A"))
                                    towar.Typ = TypTowaru.Produkt;

                                trans.Commit();
                            }
                        }
                    }

                    File.Move(fromPath + file.Name, toPath + file.Name);
                }

                session.Save();
            }
        }
    }
}
