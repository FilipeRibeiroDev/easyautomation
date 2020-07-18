# easyautomation
Library to facilitate the construction of process automation.

# Prerequisites
This release version is only available for FrameWork 4.7.2 or higher.

# Introduction
This library was created in order to facilitate the development of autonomous robots. This package includes the necessary tools for WEB automation, Windows, Image recognition including the method of solving captcha.

# Methods
## Web Automation
The web automation methods must be accessed through the **Web** attribute. Then, just inform the name of the method to be used.

## List of Types
```
TypeDriver
{
   GoogleChorme,
   InternetExplorer,
   PhantomJS,
   FireFox
}

TypeElement
{
   Id,
   Name,
   Xpath,
   CssSelector
}

TypeSelect
{
   Value,
   Text
}
```

## List of methods
### Web.StartBrowser (TypeDriver typeDriver);
Method to start browser.

### Web.CloseBrowser ();
Method to exit Browser.

### Web.Navigate (string url);
Method for navigating to a page

### Web.Click (TypeElement typeElement, string element, int timeout = 3): return void;
Method to click on element of the web page.

### Web.GetValue (TypeElement typeElement, string element, int timeout = 3): return string;
Method for getting page element value

### Web.AssignValue (TypeElement typeElement, string element, string value, int timeout = 3): return void;
Method to assign value in field

### Web.GetTableData (TypeElement typeElement, string element, int timeout = 3): return DataTable;
Method for getting data from a table

### Web.SelectValue (TypeElement typeElement, TypeSelect typeSelect, string element, string value, int timeout = 3) return void;
Method for selecting a value in a combobox

### Web.GetWebImage (TypeElement typeElement, string element, string nameImage, int timeout = 3): return Bitmap;
Method to get image from the web

### Web.ResolveCaptcha (Bitmap imageBitman): return string
Method to resolve Catpcha

# OCR automation

OCR automation is based on the image to which the robot must perform the procedure. The ideal is to print and store the exact location of the screen for the robot to use. The ocr automation methods must be accessed through the **OCR** attribute.

## List of methods

### Click (string clickImage): return void;
Method for clicking element.

### DoubleClick (string clickImage): return void;
Method to double click on element.

### DragDropClick (string clickImage, string dropImage): return void;
Method for clicking on one element based on another

# Base Automation
The base methods are used to facilitate some automation processes, in any environment. The web automation methods must be accessed through the **Base** attribute.

## List of methods
### ExtractTextPdf (string pathFile): return string;
Method to extract text from PDF

### ConvertTo <T> (IList <T> list): return DataTable;
Method to convert list to DataTable
