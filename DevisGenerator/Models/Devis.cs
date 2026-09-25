using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevisGenerator.Models
{
    class Devis
    {
        public string Chemin { get; set; }
        public string Nom { get; set; }
        public string Adresse { get; set; }
        public Devis()
        {
            this.Chemin = "./";
        }
    }
}
