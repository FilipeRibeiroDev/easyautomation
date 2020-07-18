using Sikuli4Net.sikuli_REST;
using Sikuli4Net.sikuli_UTIL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyAutomationFramework
{
    public static class OCR
    {
        //Iniciar API do Sikuli
        private static readonly APILauncher launcher = new APILauncher(true);
        /// <summary>
        /// Método para simular Click
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="clickImage"></param>
        public static void Click(string clickImage)
        {
            //Inciar API Sikuli
            launcher.Start();
            Pattern pattern = new Pattern(clickImage);

            //Simular click em OCR
            Screen scrn = new Screen();
            scrn.Wait(pattern);
            scrn.Click(pattern);

            //Encerrar API Sikuli
            launcher.Stop();
        }
        /// <summary>
        /// Método para simular Doubler Click
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="clickImage"></param>
        public static void DoubleClick(string clickImage)
        {
            //Inciar API Sikuli
            launcher.Start();
            Pattern pattern = new Pattern(clickImage);

            //Simular click em OCR
            Screen scrn = new Screen();
            scrn.Wait(pattern);
            scrn.DoubleClick(pattern);

            //Encerrar API Sikuli
            launcher.Stop();
        }
        /// <summary>
        /// Método para simular click a partir de um ponto
        /// Autor: Filipe Ribeiro
        /// </summary>
        /// <param name="clickImage"></param>
        /// <param name="dropImage"></param>
        public static void DragDropClick(string clickImage, string dropImage)
        {
            //Inciar API Sikuli
            launcher.Start();
            Pattern patternClick = new Pattern(clickImage);
            Pattern patternDrop = new Pattern(dropImage);

            //Simular click em OCR
            Screen scrn = new Screen();
            scrn.Wait(patternDrop);
            scrn.DragDrop(patternClick, patternDrop);

            //Encerrar API Sikuli
            launcher.Stop();
        }
    }
}
