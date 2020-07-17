using java.io;
using java.net;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

namespace EasyAutomationFramework
{
    public static class Base
    {
        /// <summary>
        /// Método para extrair texto de um PDF
        /// </summary>
        /// <param name="caminhoArquivo"></param>
        /// <returns></returns>
        public static string ExtractTextPdf(string caminhoArquivo)
        {
            try
            {
                URL url = new URL(caminhoArquivo);
                InputStream ist = url.openStream();
                BufferedInputStream fileToParse = new BufferedInputStream(ist);
                PDDocument document = null;
                String output = null;
                try
                {
                    document = PDDocument.load(fileToParse);
                    output = new PDFTextStripper().getText(document);
                }
                finally
                {
                    if (document != null)
                    {
                        document.close();
                    }
                    fileToParse.close();
                    ist.close();
                }
                return output;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static DataTable ConvertTo<T>(IList<T> list)
        {
            DataTable table = CreateTable<T>();
            Type entityType = typeof(T);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);
            foreach (T item in list)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item);
                }
                table.Rows.Add(row);
            }
            return table;
        }
        private static DataTable CreateTable<T>()
        {
            Type entityType = typeof(T);
            DataTable table = new DataTable(entityType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, prop.PropertyType);
            }
            return table;
        }
    }
}
