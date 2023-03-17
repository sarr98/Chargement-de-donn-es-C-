using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace app.Models
{
    public class Variable
    {
        public string mois { get; set; } = DateTime.Now.Month.ToString();
        public string annee { get; set; } = DateTime.Now.Year.ToString();
        public string idPole { get; set; }
        public string nomFichier { get; set; }
        public string CodeFamille { get; set; }
        public string NomPole { get; set; }
        //public int Periode { get; set; }
        //public int Annee { get; set; }  

       
    }
}