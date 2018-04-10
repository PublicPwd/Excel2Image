using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Excel2Image
{
    public class ExcelHelper
    {
        private string excelPath;
        private string excelDirectory;
        private int rowIndex;
        private int columnIndex;

        /// <summary>
        /// Init
        /// </summary>
        /// <param name="excelPath">The Excel document path</param>
        public ExcelHelper(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException();
            }
            if (Path.GetExtension(excelPath) != ".xlsx")
            {
                throw new Exception("Only support Microsoft Office EXCEL 2007 or higher version");
            }
            this.excelPath = excelPath;
            excelDirectory = Path.GetDirectoryName(this.excelPath) + "\\" + Path.GetFileNameWithoutExtension(this.excelPath) + "_ExcelHelper";
        }

        /// <summary>
        /// Get the original picture from the specified cell in Excel
        /// </summary>
        /// <param name="sheetName">the sheet name</param>
        /// <param name="rowIndex">the row index of the cell</param>
        /// <param name="columnIndex">the column index of the cell</param>
        /// <param name="savePath">image saving path</param>
        public void GetImage(string sheetName, int rowIndex, int columnIndex, string savePath)
        {
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;

            File.Copy(excelPath, excelPath + "_ExcelHelper");
            new FastZip().ExtractZip(excelPath + "_ExcelHelper", excelDirectory, string.Empty);
            File.Delete(excelPath + "_ExcelHelper");

            string targetSheet = GetTargetWorksheet(sheetName);
            string targetDrawing = GetTargetDrawing(targetSheet);
            string targetImage = GetTargetImage(targetDrawing);

            using (Image image = Image.FromFile(excelDirectory + @"\xl\media\" + targetImage))
            {
                image.Save(savePath);
            }

            Directory.Delete(excelDirectory, true);
        }

        private string GetTargetWorksheet(string sheetName)
        {
            string workbookXmlPath = excelDirectory + @"\xl\workbook.xml";
            XElement workbookXml = XElement.Load(workbookXmlPath);
            var sheets = workbookXml.Descendants(workbookXml.GetDefaultNamespace() + "sheet");
            XNamespace rXNamespace = workbookXml.GetNamespaceOfPrefix("r");
            var ids = from s in sheets
                      where (string)s.Attribute("name") == sheetName
                      select (string)s.Attribute(rXNamespace + "id");
            if (ids.Count() == 0)
            {
                Directory.Delete(excelDirectory, true);
                throw new Exception(sheetName + " does not exist");
            }
            string id = ids.First().ToString();

            string workbookXmlRelsPath = excelDirectory + @"\xl\_rels\workbook.xml.rels";
            XElement workbookXmlRels = XElement.Load(workbookXmlRelsPath);
            var relationships = workbookXmlRels.Elements(workbookXmlRels.GetDefaultNamespace() + "Relationship");
            string targetSheet = (from r in relationships
                                  where (string)r.Attribute("Id") == id
                                  select (string)r.Attribute("Target")).First().ToString();
            return Path.GetFileName(targetSheet);
        }

        private string GetTargetDrawing(string targetSheet)
        {
            string worksheetXmlPath = excelDirectory + @"\xl\worksheets\" + targetSheet;
            XElement worksheetXml = XElement.Load(worksheetXmlPath);
            XElement drawings = worksheetXml.Elements(worksheetXml.GetDefaultNamespace() + "drawing").First();
            XNamespace rXNamespace = worksheetXml.GetNamespaceOfPrefix("r");
            string id = (string)drawings.Attribute(rXNamespace + "id");

            string worksheetXmlRelsPath = excelDirectory + $@"\xl\worksheets\_rels\{targetSheet}.rels";
            XElement worksheetXmlRels = XElement.Load(worksheetXmlRelsPath);
            var relationships = worksheetXmlRels.Elements(worksheetXmlRels.GetDefaultNamespace() + "Relationship");
            string targetDrawing = (from r in relationships
                                    where (string)r.Attribute("Id") == id
                                    select (string)r.Attribute("Target")).First().ToString();
            return Path.GetFileName(targetDrawing);
        }

        private string GetTargetImage(string targetDrawing)
        {
            string drawingXmlPath = excelDirectory + @"\xl\drawings\" + targetDrawing;
            XElement drawingXml = XElement.Load(drawingXmlPath);
            var twoCellAnchors = drawingXml.Elements(drawingXml.GetNamespaceOfPrefix("xdr") + "twoCellAnchor");

            string id = string.Empty;
            foreach (var twoCellAnchor in twoCellAnchors)
            {
                if (!BiggerThanMinimum(twoCellAnchor.Elements().ToList()[0]))
                {
                    continue;
                }
                if (!SmallerThanMaximum(twoCellAnchor.Elements().ToList()[1]))
                {
                    continue;
                }
                id = GetTargetImage(twoCellAnchor.Elements().ToList()[2]);
                break;
            }

            if (id.Length == 0)
            {
                Directory.Delete(excelDirectory, true);
                throw new Exception("No pictures");
            }
            string drawingXmlRelsPath = excelDirectory + $@"\xl\drawings\_rels\{targetDrawing}.rels";
            XElement drawingXmlRels = XElement.Load(drawingXmlRelsPath);
            var Relationships = drawingXmlRels.Elements(drawingXmlRels.GetDefaultNamespace() + "Relationship");
            string targetImage = (from r in Relationships
                                  where (string)r.Attribute("Id") == id
                                  select (string)r.Attribute("Target")).First().ToString();
            targetImage = Path.GetFileName(targetImage);
            return targetImage;
        }

        private bool BiggerThanMinimum(XElement xElement)
        {
            string col = xElement.Elements(xElement.GetNamespaceOfPrefix("xdr") + "col").First().Value;
            string row = xElement.Elements(xElement.GetNamespaceOfPrefix("xdr") + "row").First().Value;
            return columnIndex >= int.Parse(col) && rowIndex >= int.Parse(row);
        }

        private bool SmallerThanMaximum(XElement xElement)
        {
            string col = xElement.Elements(xElement.GetNamespaceOfPrefix("xdr") + "col").First().Value;
            string row = xElement.Elements(xElement.GetNamespaceOfPrefix("xdr") + "row").First().Value;
            return columnIndex <= int.Parse(col) && rowIndex <= int.Parse(row);
        }

        private string GetTargetImage(XElement xElement)
        {
            XElement blip = xElement.Descendants(xElement.GetNamespaceOfPrefix("a") + "blip").First();
            return (string)blip.Attribute(blip.GetNamespaceOfPrefix("r") + "embed");
        }
    }
}
