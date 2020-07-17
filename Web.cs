using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using Tesseract;

namespace EasyAutomationFramework
{
    public static class Web
    {
        private static IWebDriver driver;

        /// <summary>
        /// Método para iniciar um Browser
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeDriver"></param>
        /// <returns></returns>
        public static IWebDriver StartBrowser(TypeDriver typeDriver = TypeDriver.GoogleChorme)
        {
            try
            {
                switch (typeDriver)
                {
                    case TypeDriver.GoogleChorme:
                        var sc = ChromeDriverService.CreateDefaultService();
                        sc.HideCommandPromptWindow = true;
                        ChromeOptions c = new ChromeOptions();
                        driver = new ChromeDriver(sc, c);
                        return driver;
                    case TypeDriver.PhantomJS:
                        var ps = PhantomJSDriverService.CreateDefaultService();
                        ps.HideCommandPromptWindow = true;
                        ps.AddArgument("--ignore-ssl-errors=true");
                        ps.AddArgument("--script-language='javascript'");
                        PhantomJSOptions p = new PhantomJSOptions();
                        driver = new PhantomJSDriver(ps, p);
                        return driver;
                    case TypeDriver.InternetExplorer:
                        var ie = InternetExplorerDriverService.CreateDefaultService();
                        ie.HideCommandPromptWindow = true;
                        InternetExplorerOptions i = new InternetExplorerOptions();
                        driver = new InternetExplorerDriver(ie, i);
                        return driver;
                    case TypeDriver.FireFox:
                        var fx = FirefoxDriverService.CreateDefaultService();
                        fx.HideCommandPromptWindow = true;
                        driver = new FirefoxDriver(fx);
                        return driver;
                    default:
                        return null;
                }

            }
            catch (Exception)
            {

                throw;
            }

        }
        /// <summary>
        /// Método para encerrar Browser
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="driver"></param>
        public static void CloseBrowser()
        {
            if (driver == null) return;
            try
            {
                driver.Close();
                driver.Quit();
                driver.Dispose();
            }
            catch
            {

            }
        }
        /// <summary>
        /// Método para navegar
        /// </summary>
        /// <param name="url"></param>
        public static void Navigate(string url)
        {
            driver.Navigate().GoToUrl(url);
        }
        /// <summary>
        /// Método para simular Click
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public static void Click(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement input = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                input.Click();
            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found.");
            }
        }
        /// <summary>
        /// Método para obter valor
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public static string GetValue(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement input = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                return input.Text;
            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found.");
            }
        }
        /// <summary>
        /// Método para atribuir valor em input
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="value"></param>
        /// <param name="timeout"></param>
        public static void AssignValue(TypeElement typeElement, string element, string value, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement input = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        input = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                input.SendKeys(value);
            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found");
            }

        }
        /// <summary>
        /// Obter dados da tabela
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public static DataTable GetTableData(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement webElement = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                IList<IWebElement> tableRow = webElement.FindElements(By.TagName("tr"));
                IList<IWebElement> rowTH;
                IList<IWebElement> rowTD;

                DataTable dataTable = new DataTable();
                int contador = 0;
                foreach (IWebElement row in tableRow)
                {
                    //Obter colunas
                    if (contador == 0)
                    {
                        try
                        {
                            rowTH = row.FindElements(By.TagName("th"));
                            for (int i = 0; i < rowTH.Count; i++)
                            {
                                dataTable.Columns.Add(rowTH[i].Text);
                            }
                        }
                        catch
                        {
                        }
                    }

                    rowTD = row.FindElements(By.TagName("td"));
                    if (rowTD.Count > 0)
                    {
                        List<string> myList = new List<string>();
                        for (int i = 0; i < rowTD.Count; i++)
                        {
                            myList.Add(rowTD[i].Text);
                        }
                        dataTable.Rows.Add(myList.ToArray());

                    }
                    contador++;
                }


                return dataTable;
            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found");
            }
        }
        /// <summary>
        /// Método para selecionar valor em combobox
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="typeSelect"></param>
        /// <param name="element"></param>
        /// <param name="value"></param>
        /// <param name="timeout"></param>
        public static void SelectValue(TypeElement typeElement, TypeSelect typeSelect, string element, string value, int timeout = 3)
        {
            try
            {

                SelectElement select = null;
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));

                switch (typeElement)
                {
                    case TypeElement.Id:
                        select = new SelectElement(wait.Until(ExpectedConditions.ElementExists(By.Id(element))));
                        break;
                    case TypeElement.Name:
                        select = new SelectElement(wait.Until(ExpectedConditions.ElementExists(By.Name(element))));
                        break;
                    case TypeElement.Xpath:
                        select = new SelectElement(wait.Until(ExpectedConditions.ElementExists(By.XPath(element))));
                        break;
                    case TypeElement.CssSelector:
                        select = new SelectElement(wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element))));
                        break;
                }

                switch (typeSelect)
                {
                    case TypeSelect.Text:
                        select.SelectByText(value);
                        break;
                    case TypeSelect.Value:
                        select.SelectByValue(value);
                        break;
                }

            }
            catch (Exception)
            {
                throw new Exception($"Option {value} could not be selected on element {element}.");
            }
        }
        /// <summary>
        /// Método para obter imagem da web
        /// Autor: Filipe Ribeiro 
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="uniqueName"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public static Image GetWebImage(TypeElement typeElement, string element, string nameImage, int timeout = 3)
        {
            try
            {
                Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                screenshot.SaveAsFile(Directory.GetCurrentDirectory() + nameImage, ScreenshotImageFormat.Png);

                Image img = Bitmap.FromFile(Directory.GetCurrentDirectory() + nameImage);
                Rectangle rect = new Rectangle();
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement webElement = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                if (webElement != null)
                {
                    // Get the Width and Height of the WebElement using
                    int width = webElement.Size.Width;
                    int height = webElement.Size.Height;

                    // Get the Location of WebElement in a Point.
                    // This will provide X & Y co-ordinates of the WebElement
                    Point p = webElement.Location;

                    // Create a rectangle using Width, Height and element location
                    rect = new Rectangle(p.X, p.Y, width, height);
                }

                // croping the image based on rect.
                Bitmap bmpImage = new Bitmap(img);
                var cropedImag = bmpImage.Clone(rect, bmpImage.PixelFormat);

                return cropedImag;
            }
            catch (Exception)
            {

                throw new Exception($"Element {element} not found");
            }
        }
        /// <summary>
        /// Método para resolver captcha
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="imagePix"></param>
        /// <returns></returns>
        public static string ResolveCaptcha(Pix imagePix)
        {
            string res = "";
            try
            {

                using (var engine = new TesseractEngine("tessdata", "eng", EngineMode.Default))
                {
                    engine.SetVariable("tessedit_char_whitelist", "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
                    engine.SetVariable("tessedit_unrej_any_wd", true);

                    using (var page = engine.Process(imagePix, PageSegMode.Auto))
                    {
                        res = page.GetText();
                    }
                }
            }
            catch
            {

            }

            return res.Replace("\n", "").Replace(" ", "").Replace(" ", "");
        }
        /// <summary>
        /// Método para acessar Pop Up
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public static void AccessPopUpClick(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement webElement = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                PopupWindowFinder finder = new PopupWindowFinder(driver);
                string popupWindowHandle = finder.Click(webElement);
                driver.SwitchTo().Window(popupWindowHandle);

            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found");
            }
        }
        /// <summary>
        /// Método para sair do Pop Up
        /// Autor: Filipe Ribeiro
        /// </summary>
        public static void LeavePopUp()
        {
            try
            {
                driver.SwitchTo().DefaultContent();
            }
            catch (Exception)
            {
                throw new Exception($"Driver not found");
            }
        }
        /// <summary>
        /// Método para acessar iframe
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public static void AccessFrame(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement webElement = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }

                driver.SwitchTo().Frame(webElement);
            }
            catch (Exception)
            {
                throw new Exception($"Element {element} not found");
            }
        }
        /// <summary>
        /// Método para executar script
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="script"></param>
        public static void ExecuteScript(string script)
        {
            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript(script);
            }
            catch (Exception)
            {
                throw new Exception($"Falha Script");
            }
        }
        /// <summary>
        /// Verificar se página carregou
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="timeout"></param>
        public static void WaitForLoad(int timeout = 60)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                wait.Until(wd => ((IJavaScriptExecutor)wd).ExecuteScript("return document.readyState").ToString() == "complete");
            }
            catch
            {
            }
        }
        /// <summary>
        /// Método para verificar se elemento existe na tela
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public static void WaitForElement(TypeElement typeElement, string element, int timeout = 3)
        {
            try
            {

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement webElement = null;
                switch (typeElement)
                {
                    case TypeElement.Id:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Name:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Id(element)));
                        break;
                }


                wait.Until(ElementIsVisible(webElement));
            }
            catch (Exception)
            {

                throw;
            }
        }
        private static Func<IWebDriver, bool> ElementIsVisible(IWebElement element)
        {

            return (_dr) =>
            {
                try
                {
                    return element.Displayed;
                }
                catch (Exception)
                {
                    // If element is null, stale or if it cannot be located
                    return false;
                }
            };
        }
    }
    /// <summary>
    /// Tipos de Driver
    /// </summary>
    public enum TypeDriver
    {
        GoogleChorme,
        InternetExplorer,
        PhantomJS,
        FireFox
    }
    /// <summary>
    /// Tipos de Elementos
    /// </summary>
    public enum TypeElement
    {
        Id,
        Name,
        Xpath,
        CssSelector
    }
    /// <summary>
    /// Tipo de seleção
    /// </summary>
    public enum TypeSelect
    {
        Value,
        Text
    }
}
