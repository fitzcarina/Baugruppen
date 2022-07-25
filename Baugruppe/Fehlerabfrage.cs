using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Baugruppe
{
    internal class Fehlerabfrage
    {

        public static void fehlerabfrage_autragsnummer(string auftragsnummer)
        {

            if (auftragsnummer.Length < 1 || auftragsnummer.Length >10)
            {
                MessageBox.Show("Bitte geben Sie eine Auftagsnummer ein !");
            }
          
        }



    }
}
