using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcelToVCF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog();

        }

        private void OpenFileDialog()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                Filter = "Arquivos excel|*.xlsx",
                Title = "Selecione o arquivo"
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var contacts = ExtractContact(openFileDialog1.FileName);
                var fileString = WriteVCFFile(contacts);
                SaveFile(fileString);
                Close();
            }
        }

        private List<Contact> ExtractContact(string path)
        {
            var contacts = new List<Contact>();


            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            try
                            {
                                var phone = reader.GetValue(0).ToString();
                                var name = reader.GetValue(1).ToString();
                                contacts.Add(new Contact()
                                {
                                    Name = name,
                                    Phone = phone
                                });
                            }
                            catch
                            {

                            }
                        }
                    } while (reader.NextResult());
                }
            }
            
            //Removo os nomes das colunas
            contacts.RemoveAt(0);
            return contacts; 
        }

        private string WriteVCFFile(List<Contact> contacts)
        {
            var fileString = new StringBuilder();

            foreach (var contact in contacts)
            {
                fileString.AppendFormat("BEGIN: VCARD{0}", Environment.NewLine);
                fileString.AppendFormat("VERSION:2.1{0}", Environment.NewLine);
                fileString.AppendFormat("N: {0}; ; ; ;{1}", contact.Name, Environment.NewLine);
                fileString.AppendFormat("FN: {0}{1}", contact.Name, Environment.NewLine);
                fileString.AppendFormat("SOUND; X - IRMC - N:; ; ; ;{0}", Environment.NewLine);
                fileString.AppendFormat("TEL; CELL: {0}{1}", contact.Phone, Environment.NewLine);
                fileString.AppendFormat("END:VCARD{0}", Environment.NewLine);
            }

            return fileString.ToString();
        }

        private void SaveFile(string fileStr)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Arquivos VCF|*.vcf";
            saveFileDialog1.Title = "Salve o arquivo VCF";
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                File.WriteAllText(saveFileDialog1.FileName, fileStr);
            }
        }

    }

    public class Contact
    {
        public string Phone { get; set; }


        public string Name { get; set; }

    }
}
