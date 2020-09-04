using System;
using System.Data;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Xsl;
using System.IO;
using System.Text;

namespace XMLToWordApp
{
    public partial class Form1 : Form
    {
        private DataSet ds;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //Create the DataSet from the XML file
                ds = new DataSet();
                ds.ReadXml("XSLFile.xml");
                var a = ds.GetXml();
                GenerateDoc("Wololo", "XSLFile.xslt", "Output.doc", a);

       
                MessageBox.Show("Document created successfully.....");
            }
            catch (Exception ex)
            {
                ds = null;
                MessageBox.Show(ex.StackTrace);
            }
        }

        /// <summary>
        /// Генерит word документ на основе xsl шаблона в указанный фолдер
        /// </summary>
        /// <param name="archiveFolder"></param>
        /// <param name="xslFile"></param>
        /// <param name="fileName"></param>
        /// <param name="xml"></param>
        private void GenerateDoc(string archiveFolder, string xslFile, string fileName, string xml)
        {
            try
            {
                var xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);

                var transformedXmlDoc = XslTransform(Alphaleonis.Win32.Filesystem.Path.Combine("", xslFile), xmlDoc);
                var outputStr = transformedXmlDoc.Replace("</table>", "</table><br/>");
                outputStr = $"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><style type=\"text/css\"></style>{"<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]--><!--@page Section1   {size:8.5in 11.0in;  margin:1.0in 1.25in 1.0in 1.25in ;   mso-header-margin:.5in;    mso-footer-margin:.5in; mso-paper-source:0;} div.Section1{page:Section1;}-->" + "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" + "</head><body>"}" + outputStr + "</body></html>";
                Alphaleonis.Win32.Filesystem.Directory.CreateDirectory(archiveFolder);
                var file = Alphaleonis.Win32.Filesystem.File.CreateText(archiveFolder + "/" + fileName);
                file.Write(outputStr);
                file.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Преобразует XML документ в соответствии со XSL схемой
        /// </summary>
        /// <param name="xslPath">путь к XSL схеме</param>
        /// <param name="data">XML документ</param>
        /// <returns></returns>
        public static string XslTransform(string xslPath, XmlDocument data)
        {
            XslCompiledTransform xslTransform = new XslCompiledTransform();

            string path = xslPath;

            XmlDocument xsltDoc;
            XmlReader reader;
            using (reader = XmlReader.Create(path))
            {
                xsltDoc = new XmlDocument();
                xsltDoc.Load(reader);
                reader.Close();
            }

            xslTransform.Load(xsltDoc, new XsltSettings(true, true), new XmlUrlResolver());

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(data.InnerXml);

            reader = xml.CreateNavigator().ReadSubtree();
            StringBuilder sb = new StringBuilder();
            using (StringWriter writer = new StringWriter(sb))
            {
                XsltArgumentList xsltArgs = new XsltArgumentList();
                xslTransform.Transform(reader, xsltArgs, writer);
                writer.Close();
                reader.Close();
            }

            return sb.ToString();
        }
    }
}