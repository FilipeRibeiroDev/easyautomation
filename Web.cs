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
using Tesseract;

namespace EasyAutomationFramework
{
    /// <summary>
    /// Resumo: Dar Suporte para desenvolvimento de automação Web.
    /// </summary>
    public class Web 
    {
        private IWebDriver driver;
        /// <summary>
        /// Método para iniciar um Browser
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeDriver"></param>
        /// <returns></returns>
        public EasyReturn.Web StartBrowser(TypeDriver typeDriver = TypeDriver.GoogleChorme)
        {
            try
            {
                switch (typeDriver)
                {
                    case TypeDriver.GoogleChorme:
                        var sc = ChromeDriverService.CreateDefaultService();
                        sc.HideCommandPromptWindow = true;
                        ChromeOptions c = new ChromeOptions();
                        c.AddArgument("--incognito");
                        c.AddExcludedArgument("enable-automation");
                        c.AddAdditionalCapability("useAutomationExtension", false);
                        c.AddArgument("--start-maximized");
                        driver = new ChromeDriver(sc, c);
                        break;
                    case TypeDriver.PhantomJS:
                        var ps = PhantomJSDriverService.CreateDefaultService();
                        ps.HideCommandPromptWindow = true;
                        ps.AddArgument("--ignore-ssl-errors=true");
                        ps.AddArgument("--script-language='javascript'");
                        PhantomJSOptions p = new PhantomJSOptions();
                        driver = new PhantomJSDriver(ps, p);
                        break;
                    case TypeDriver.InternetExplorer:
                        var ie = InternetExplorerDriverService.CreateDefaultService();
                        ie.HideCommandPromptWindow = true;
                        InternetExplorerOptions i = new InternetExplorerOptions();
                        driver = new InternetExplorerDriver(ie, i);
                        break;
                    case TypeDriver.FireFox:
                        var fx = FirefoxDriverService.CreateDefaultService();
                        fx.HideCommandPromptWindow = true;
                        driver = new FirefoxDriver(fx);
                        break;
                    default:
                        break;
                }

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = ex.Message.ToString()
                };
            }

        }
        /// <summary>
        /// Método para encerrar Browser
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="driver"></param>
        public void CloseBrowser()
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
        public EasyReturn.Web Navigate(string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"More info: {ex.Message}"
                };
            }
        }
        /// <summary>
        /// Método para simular Click
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public EasyReturn.Web Click(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }

                webElement.Click();
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"Element {element} not found. More info: {ex.Message}"
                };
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
        public EasyReturn.Web GetValue(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }

                return new EasyReturn.Web
                {
                    driver = driver,
                    Value = webElement.Text,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"Element {element} not found. More info: {ex.Message}"
                };
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
        public EasyReturn.Web AssignValue(TypeElement typeElement, string element, string value, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }

                webElement.SendKeys(value);
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"Element {element} not found. More info: {ex.Message}"
                };
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
        public EasyReturn.Web GetTableData(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
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

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true,
                    table = dataTable
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"Element {element} not found. More info: {ex.Message}"
                };
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
        public EasyReturn.Web SelectValue(TypeElement typeElement, TypeSelect typeSelect, string element, string value, int timeout = 3)
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
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = false,
                    Error = $"Option {value} could not be selected on element {element}. More info: {ex.Message}"
                };
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
        public EasyReturn.Web GetWebImage(TypeElement typeElement, string element, string nameImage, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
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

                return new EasyReturn.Web
                {
                    driver = driver,
                    Bitmap = (Bitmap)cropedImag,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"Element {element} not found. More info: {ex.Message}",
                    Sucesso = false
                };
            }
        }
        /// <summary>
        /// Método para resolver captcha
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="imageBitman"></param>
        /// <returns></returns>
        public EasyReturn.Web ResolveCaptcha(Bitmap imageBitman)
        {
            string res = "";
            try
            {

                using (var engine = new TesseractEngine("./tessdata", "eng", EngineMode.Default))
                {
                    engine.SetVariable("tessedit_char_whitelist", "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
                    engine.SetVariable("tessedit_unrej_any_wd", true);

                    using (var page = engine.Process(imageBitman, PageSegMode.Auto))
                    {
                        res = page.GetText();
                    }
                }

                return new EasyReturn.Web
                {
                    driver = driver,
                    Value = res.Replace("\n", "").Replace(" ", "").Replace(" ", ""),
                    Sucesso = true
                };
            }
            catch(Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = ex.Message.ToString(),
                    Sucesso = false
                };
            }

             
        }
        /// <summary>
        /// Método para acessar Pop Up
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public EasyReturn.Web AccessPopUpClick(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }

                PopupWindowFinder finder = new PopupWindowFinder(driver);
                string popupWindowHandle = finder.Click(webElement);
                driver.SwitchTo().Window(popupWindowHandle);

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"Element {element} not found. More info: {ex.Message}",
                    Sucesso = false
                };
            }
        }
        /// <summary>
        /// Método para sair do Pop Up
        /// Autor: Filipe Ribeiro
        /// </summary>
        public EasyReturn.Web LeavePopUp()
        {
            try
            {
                driver.SwitchTo().DefaultContent();
                
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"Driver not found. More info: {ex.Message}",
                    Sucesso = false
                };
            }
        }
        /// <summary>
        /// Método para acessar iframe
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="typeElement"></param>
        /// <param name="element"></param>
        /// <param name="timeout"></param>
        public EasyReturn.Web AccessFrame(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }

                driver.SwitchTo().Frame(webElement);

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"Element {element} not found. More info: {ex.Message}",
                    Sucesso = false
                };
            }
        }
        /// <summary>
        /// Método para executar script
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="script"></param>
        public EasyReturn.Web ExecuteScript(string script)
        {
            try
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript(script);

                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"Falha execute in Script. More info: {ex.Message}",
                    Sucesso = false
                };
            }
        }
        /// <summary>
        /// Verificar se página carregou
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="timeout"></param>
        public void WaitForLoad(int timeout = 60)
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
        public EasyReturn.Web WaitForElement(TypeElement typeElement, string element, int timeout = 3)
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
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.Name(element)));
                        break;
                    case TypeElement.Xpath:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.XPath(element)));
                        break;
                    case TypeElement.CssSelector:
                        webElement = wait.Until(ExpectedConditions.ElementExists(By.CssSelector(element)));
                        break;
                }


                wait.Until(ElementIsVisible(webElement));
                return new EasyReturn.Web
                {
                    driver = driver,
                    Sucesso = true
                };
            }
            catch (Exception ex)
            {
                return new EasyReturn.Web
                {
                    driver = driver,
                    Error = $"More info: {ex.Message}",
                    Sucesso = false
                };
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
