using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using MixERP.Net.VCards;
using MixERP.Net.VCards.Models;
using MixERP.Net.VCards.Serializer;
using MixERP.Net.VCards.Types;

namespace ContactTableToVcf
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
           
           
            ;
            //Application.Run(new Form1());

            Process();
        }

        static void Process()
        {
            try
            {
                CNTRecord CNTRecord = new CNTRecord();
                System.Xml.Serialization.XmlSerializer serializer = new XmlSerializer(typeof(CNTRecord));
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "打开ContactTable.xml格式通讯录";
                openFileDialog.Filter = "通讯录备份|*.xml";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var fs = openFileDialog.OpenFile())
                    {
                        try
                        {
                            CNTRecord = (CNTRecord) serializer.Deserialize(fs);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.Message);
                            return;
                        }
                    }

                }
                else
                {
                    return;
                }


                var vcardList = new List<VCard>();
                foreach (Contact contact in CNTRecord)
                {
                    var vcard = new VCard
                    {
                        Version = VCardVersion.V4,
                        FormattedName = contact.Name, // "John Doe",
                        //FirstName = "John",
                        //LastName = "Doe",
                        //Classification = ClassificationType.Confidential,
                        //Categories = new[] { "Friend", "Fella", "Amsterdam" },
                        //...
                        Note = contact.Note,

                    };


                    if (contact.EmailList != null)
                    {

                        var Emails = new List<MixERP.Net.VCards.Models.Email>();
                        foreach (Email email in contact.EmailList)
                        {
                            Emails.Add(new MixERP.Net.VCards.Models.Email()
                            {
                                EmailAddress = email.Value,
                                Type = (EmailType) int.Parse(email.Type),
                            });
                        }
                        vcard.Emails = Emails;
                    }
                    if (contact.PostalAddressList != null)
                    {
                        var Addresses = new List<Address>();
                        foreach (var item in contact.PostalAddressList)
                        {
                            Addresses.Add(new MixERP.Net.VCards.Models.Address()
                            {
                                ExtendedAddress = item.Value,
                                Type = (AddressType) int.Parse(item.Type),
                            });
                        }
                        vcard.Addresses = Addresses;
                    }
                    if (contact.PhoneList != null)
                    {
                        var Telephones = new List<Telephone>();
                        foreach (var item in contact.PhoneList)
                        {
                            Telephones.Add(new MixERP.Net.VCards.Models.Telephone()
                            {
                                Number = item.Value,
                                Type = (TelephoneType) int.Parse(item.Type),
                            });
                        }
                        vcard.Telephones = Telephones;
                    }
                    vcardList.Add(vcard);
                }
                string cvfs = vcardList.Serialize();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "保存到";
                saveFileDialog.Filter = "vcf文件|*.vcf";
                saveFileDialog.FileName = DateTime.Now.ToString("yyyy-MM-dd") + ".vcf";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var fs = saveFileDialog.OpenFile())
                    {
                        TextWriter textWriter = new StreamWriter(fs);
                        textWriter.Write(cvfs);
                        textWriter.Dispose();
                    }

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }
    }

    [XmlRoot("CNTRecord")]
    public class CNTRecord: List<Contact>
    {
       // public List<Contact> CNTRecord { get; set; }
    }

    public class Contact
    {
        public string Name { get; set; }

        public string Starred { get; set; }


        public List<Phone> PhoneList { get; set; }


        public string Note { get; set; }

        public List<Organization> OrganizationList { get; set; }

        public Account Account { get; set; }

        public List<PostalAddress> PostalAddressList { get; set; }


        public List<Email> EmailList { get; set; }

        /// <summary>
        /// base64
        /// </summary>
        public String HeadSculpture { get; set; }


    }


    public class Phone:System.Xml.Serialization.IXmlSerializable
    {
        public string Type { get; set; }

        public string Value { get; set; }

        public XmlSchema GetSchema()
        {
            return new XmlSchema();
        }

        public void ReadXml(XmlReader reader)
        {
           // reader.MoveToAttribute("type");
            this.Type = reader.GetAttribute("Type");
            this.Value = reader.ReadElementContentAsString();
        }

        public void WriteXml(XmlWriter writer)
        {
           throw new NotImplementedException();
        }
    }


    public class Organization
    {
        [XmlAttribute]
        public string Type { get; set; }

        public Detail Detail { get; set; }




    }

    public class Detail
    {
        public string Company { get; set; }
    }


    public class Account
    {

        [XmlAttribute]
        public string Value { get; set; }


        public string Name { get; set; }

        public string Type { get; set; }

    }



    public class PostalAddress : System.Xml.Serialization.IXmlSerializable
    {
        public string Type { get; set; }

        public string Value { get; set; }

        public XmlSchema GetSchema()
        {
            return new XmlSchema();
        }

        public void ReadXml(XmlReader reader)
        {
            // reader.MoveToAttribute("type");
            this.Type = reader.GetAttribute("Type");
            this.Value = reader.ReadElementContentAsString();
        }

        public void WriteXml(XmlWriter writer)
        {
            throw new NotImplementedException();
        }
    }


    public class Email : System.Xml.Serialization.IXmlSerializable
    {
        public string Type { get; set; }

        public string Value { get; set; }

        public XmlSchema GetSchema()
        {
            return new XmlSchema();
        }

        public void ReadXml(XmlReader reader)
        {
            // reader.MoveToAttribute("type");
            this.Type = reader.GetAttribute("Type");
            this.Value = reader.ReadElementContentAsString();
        }

        public void WriteXml(XmlWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
