using HtmlAgilityPack;
using ScrapySharp.Extensions;
using ScrapySharp.Html;
using ScrapySharp.Html.Forms;
using ScrapySharp.Network;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.UI.WebControls.WebParts;

namespace ConsoleApplication1
{
    class Program
    {
        static int Main(string[] args)
        {
            string delay;
            string logPath;
            int min = 0;
            int max = 0;

            if (args.Length > 5 || args.Length < 3)
            {
                Console.WriteLine("Numero de parametros invalido");
                return 1;
            }
            bool resultMin = int.TryParse(args[1], out min);
            bool resultMax = int.TryParse(args[2], out max);

            if (resultMin == true && resultMax == true)
            {
                if (min > max)
                {
                    Console.WriteLine("Rango invalido");
                    return 1;
                }
            }else
            {
                Console.WriteLine("Rango invalido");
                return 1;
            }
            if (args.Length == 3)
            {
                delay = "10";
                logPath = "log.txt";
            }
            else
            {
                if (args.Length == 4)
                {
                    int i = 0;
                    bool result = int.TryParse(args[3], out i);
                    if (result == true)
                    {
                        delay = args[3];
                        logPath = "log.txt";
                    }
                    else
                    {
                        delay = "10";
                        logPath = args[3];
                    }
                }
                else
                {
                    int i = 0;
                    bool result = int.TryParse(args[3], out i);
                    if (result == true)
                    {
                        delay = args[3];
                        logPath = args[4];
                    } else
                    {
                        delay = args[4];
                        logPath = args[3];
                    }
                }
            }
            using (System.IO.StreamWriter log = new StreamWriter(@logPath, true))
            {
                bool existe = false;
                if (File.Exists(args[0]))
                {
                    existe = true;
                }
                using (System.IO.StreamWriter fs = new StreamWriter(@args[0], true, System.Text.Encoding.ASCII))
                {
                    if (!existe)
                    {
                        string[] encabezado = new string[] { "Matricula;", "Nombres;", "Documento;", "Cuit;", "Ramo;", "Domicilio;", "Localidad;",
                        "Provincia;", "Telefonos;", "Email;", "Cod. Postal;", "Info;"};
                        foreach (string str in encabezado)
                        {
                            fs.Write(str);
                        }
                        fs.WriteLine("");
                    }
                    string[] listaTexto = new string[] { "Matr&iacute;cula:", "Nombres:", "Documento:", "Cuit:", "Ramo:", "Domicilio:", "Localidad:",
                    "Provincia:", "Tel&eacute;fonos:", "Email:"};
                    string rangoMinString = "";
                    string rangoMaxString = "";
                    string replacement;

                    rangoMinString = args[1];
                    rangoMaxString = args[2];
                    int rangoMin = Int32.Parse(rangoMinString);
                    int rangoMax = Int32.Parse(rangoMaxString);
                    for (int i = rangoMin; i <= rangoMax; i++)
                    {
                        log.Write(DateTime.Now);
                        ScrapingBrowser browser = new ScrapingBrowser();
                        WebPage homePage;
                        try
                        {
                            homePage = browser.NavigateToPage(new Uri("http://www.ssn.gov.ar/storage/registros/productores/productoresactivosfiltro.asp"));
                        }
                        catch (Exception e)
                        {
                            log.WriteLine(" - NO SE PUDO CONECTAR CON LA PAGINA ");
                            return 1;
                        }
                        log.Write(" - CONEXION CON LA PAGINA: OK ");
                        PageWebForm form = homePage.FindForm("form1");
                        form["matricula"] = i.ToString();
                        WebPage resultsPage;
                        try
                        {
                            resultsPage = form.Submit(new Uri("http://www.ssn.gov.ar/storage/registros/productores/productoresactivos.asp"), HttpVerb.Post);
                        }
                        catch (Exception e)
                        {
                            log.WriteLine(" - POST MATRICULA [" + i.ToString() + "] : FALLO ");
                            return 1;
                        }
                        log.Write(" - POST MATRICULA [" + i.ToString() + "] : OK ");
                        var html = new HtmlDocument();
                        html.LoadHtml(resultsPage.Content.ToString());
                        HtmlNode nodeTest = html.DocumentNode.SelectSingleNode("//font[contains(text(),'Sin Registros para los filtros aplicados')]");
                        if (nodeTest == null)
                        {
                            foreach (string str in listaTexto)
                            {
                                HtmlNode tr = html.DocumentNode.SelectSingleNode(string.Format("//font[contains(text(),'{0}')]/ancestor::tr", str));
                                HtmlNode font = tr.SelectSingleNode("descendant::font[@color]");
                                replacement = Regex.Replace(font.InnerText, @"\t|\n|\r|  |&nbsp|;", "");
                                fs.Write(replacement);
                                fs.Write(";");
                            }
                            HtmlNode postal = html.DocumentNode.SelectSingleNode("//font[contains(text(),'Postal:')]");
                            replacement = Regex.Replace(postal.InnerText, @"\t|\n|\r|  |&nbsp|;", "");
                            fs.Write(replacement.Substring(12).Trim());
                            fs.Write(";");
                            fs.WriteLine("OK");
                            log.Write("- RESPUESTA: OK");
                            log.WriteLine("");
                        }
                        else
                        {
                            log.Write("- RESPUESTA: NO ENCONTRADO");
                            log.WriteLine("");
                            fs.WriteLine(i.ToString()+";;;;;;;;;;; No se encuentran datos");
                        }
                        Thread.Sleep(Int32.Parse(delay) * 1000);
                    }
                    return 0;
                }
            }
        }
    }
}