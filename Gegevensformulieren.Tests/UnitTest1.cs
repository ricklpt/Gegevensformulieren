using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Gegevensformulieren.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestInitialize]
        public void Initialize()
        {
            var file = File.ReadAllText("c:\\test.xml");
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(file);

            var nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
            var root = xmlDoc.DocumentElement;

            var node = xmlDoc.SelectNodes("//ss:Row", nsmgr);

            var positions = new Dictionary<int, string>();
            var persons = new List<Persoon>();
            for (var i = 0; i < node.Count; i++)
            {
                var item = node[i];
                var dataNodes = item.SelectNodes("ss:Cell/ss:Data", nsmgr);

                if (i == 0)
                {
                    //GetPositions
                    for (var j = 0; j < dataNodes.Count; j++)
                    {
                        positions.Add(j, dataNodes[j].InnerText);
                    }
                }
                else
                {
                    persons.Add(new Persoon(dataNodes, positions));
                }
            }
        }

        [TestMethod]
        public void TestMethod1()
        {
        }
    }

    public class Persoon
    {
        public Persoon(XmlNodeList nodesList, Dictionary<int, string> positions)
        {
            foreach (var prop in typeof (Persoon).GetProperties())
            {
                var attrs = (NameAttribute[]) prop.GetCustomAttributes
                    (typeof (NameAttribute), false).First();
                foreach (var attr in attrs)
                {
                    prop.SetValue(this,
                        nodesList[
                            positions.First(
                                x => string.Compare(x.Value, attr.Name, StringComparison.OrdinalIgnoreCase) == 0).Key]
                            .InnerText);
                }
            }
        }

        [Name("lidnummer")]
        public string Lidnummer { get; set; }

        [Name("lid initialen")]
        public string Initialen { get; set; }

        [Name("lid voornaam")]
        public string Voornaam { get; set; }

        [Name("lid tussenvoegsel")]
        public string Tussenvoegsel { get; set; }

        [Name("lid achternaam")]
        public string Achternaam { get; set; }

        [Name("lid geslacht")]
        public string Geslacht { get; set; }

        [Name("lid straat")]
        public string Straat { get; set; }

        [Name("lid huisnummer")]
        public string Huisnummer { get; set; }

        [Name("Lid toevoegsel huisnr")]
        public string Huisnummertoevoegsel { get; set; }

        [Name("lid postcode")]
        public string Postcode { get; set; }

        [Name("lid plaats")]
        public string Plaats { get; set; }

        [Name("land")]
        public string Land { get; set; }

        [Name("lid geboortedatum")]
        public string Geboortedatum { get; set; }

        [Name("lid geboorteplaats")]
        public string Geboorteplaats { get; set; }

        [Name("lid mailadres")]
        public string Mailadres { get; set; }

        [Name("lid telefoon")]
        public string Telefoon { get; set; }

        [Name("lid mobiel")]
        public string Mobiel { get; set; }

        [Name("organisatienummer")]
        public string Organisatienummer { get; set; }

        [Name("organisatie")]
        public string Organisatie { get; set; }

        [Name("plaats")]
        public string plaats { get; set; }

        [Name("speleenheid")]
        public string Speleenheid { get; set; }

        [Name("deelnemer functie")]
        public string Functie { get; set; }

        [Name("Lid mailadres ouder/verzorger 1")]
        public string MailAdresOuderEen { get; set; }

        [Name("Lid mailadres ouder/verzorger 2")]
        public string MailAdresOuderTwee { get; set; }

        [Name("De gegevens van mijn kind kloppen:")]
        public string GegevensKloppen { get; set; }

        [Name("Het emailadres ouders/verzorgers is ingevuld en klopt:")]
        public string EmailKlopt { get; set; }

        [Name("Telefoonnummer ouders:")]
        public string TelefoonOuders { get; set; }

        [Name("Mobiel telefoonnummer ouders:")]
        public string MobielOuderEen { get; set; }

        [Name("Naam:")]
        public string Naam { get; set; }

        [Name("Adres:")]
        public string Adres { get; set; }

        [Name("Plaats:")]
        public string PlaatsOuderEen { get; set; }

        [Name("Telefoonnummer:")]
        public string Telefoonnummer { get; set; }

        [Name("Relatie tot uw kind:")]
        public string Relatietotuwkind { get; set; }

        [Name("Heeft uw kind een zwemdiploma?: Nee")]
        public string GeenZwemDiploma { get; set; }

        [Name("Heeft uw kind een zwemdiploma?: A")]
        public string ZwemDiplomaA { get; set; }

        [Name("Heeft uw kind een zwemdiploma?: B")]
        public string ZwemDiplomaB { get; set; }

        [Name("Heeft uw kind een zwemdiploma?: C")]
        public string ZwemDiplomaC { get; set; }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class NameAttribute : Attribute
    {
        public string Name;

        public NameAttribute(string name)
        {
            Name = name;
        }
    }
}