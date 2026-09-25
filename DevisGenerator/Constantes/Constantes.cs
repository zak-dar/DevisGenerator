using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevisGenerator.Constantes
{
    public static class Constantes
    {
        public static readonly string FREQUENCE = "Fréquence : {0}";

        public static readonly string NETTOYAGE_MARCHES = "  - Nettoyage des marches et contremarches d'accès au hall d'entrée\r\n";
        public static readonly string NETTOYAGE_CUIVRES = "  - Nettoyage des cuivres\r\n";
        public static readonly string NETTOYAGE_PORTE_TAMBOUR = "  - Nettoyage de la porte à tambour\r\n";
        public static readonly string NETTOYAGE_PORTES_ASCENCEUR = "  - Nettoyage des portes d'ascenseur de chaque palier";
        public static readonly string NETTOYAGE_RAILS_ASCENCEUR = "  - Nettoyage des rails d'ascenseur de chaque palier";
        public static readonly string ENLEVEMENT_DETRITUTS_FEUILLES_VOIE_ACCES_GARAGE = "  - Enlèvement des détritus et des feuilles de la voie d'accès aux garages";
        public static readonly string SORTIE_CONTENEUR = "  - Sortie des conteneurs";
        public static readonly string SORTIE_SACS = "  - Sortie des sacs";
        public static readonly string RENTREE_CONTENEURS = "\r\n  - Rentrée des conteneurs";

        public static readonly string TITRE_SECTION_CAVES_ET_ACCES_CAVES = "Caves et accès caves\r\n";
        public static readonly string TITRE_SECTION_NETTOYAGE_CABINE_ASCENCEUR = "Nettoyage de la cabine d'ascenseur\r\n";
        public static readonly string TITRE_SECTION_GARAGE = "Garage\r\n";
        public static readonly string TITRE_SECTION_ACCES_GARAGE = "Accès garages\r\n";
        public static readonly string TITRE_SECTION_ORDURES_MENAGERES = "Ordures Ménagères\r\n";
        public static readonly string TITRE_SECTION_ABORDS_ACCES = "Abords & accès\r\n";
        public static readonly string TITRE_SECTION_LOCAL_CONTENEUR_VIDE_ORDURE = "Local conteneur / vide ordure\r\n";

        public static readonly string SECTION_ACCES_GARAGE = "  - Balayage et lavage des marches et contremarches des accès aux garages\r\n  - Dépoussiérage des plinthes\r\n  - Enlèvement des toiles d'araignée\r\n  - Balayage des sols des sas d'accès au garage\r\n  - Lavage des sols à l'aide d'un produit détergent et odorant\r\n  - Nettoyage des boutons de minuterie\r\n  - Dépoussiérage des luminaires";
        public static readonly string SECTION_CAVES_ET_ACCES_CAVES = "  - Enlèvement des toiles d'araignée\r\n  - Nettoyage des boutons de minuterie\r\n  - Dépoussiérage des luminaires\r\n  - Balayage et lavage des couloirs de caves";
        public static readonly string SECTION_NETTOYAGE_CABINE_ASCENCEUR = "  - Nettoyage des sols\r\n  - Nettoyage des parois\r\n  - Nettoyage des plaques de commande\r\n  - Nettoyage des rails d'ascenseur\r\n  - Nettoyage des miroirs\r\n  - Nettoyage des portes d'ascenseur";
        public static readonly string SECTION_GARAGE = "  - Enlèvement des détritus\r\n  - Enlèvement des éventuels détritus dans les bacs à sable pour la protection incendie\r\n  - Dépoussiérage des extincteurs\r\n  - Enlèvement des toiles d'araignée\r\n  - Balayage des bordures des voies de circulation";
        public static readonly string SECTION_ABORDS_ACCES = "  - Enlèvement des feuilles mortes aux abords de la copropriété\r\n  - Enlèvement des détritus aux abords de la copropriété\r\n  - Balayage des abords de la copropriété\r\n";        
        public static readonly string SECTION_LOCAL_CONTENEUR_VIDE_ORDURE = "  - Enlèvement des détritus\r\n  - Balayage des sols du local \r\n  - Lavage des sols avec un produit bactéricide, fongicide et surodorant\r\n  - Lavage des conteneurs avec un produit bactéricide, fongicide et surodorant\r\n  - Enlèvement des toiles d'araignée";
        public static readonly string SECTION_ORDURES_MENAGERES = "  - Enlèvement des détritus au sol";
        public static readonly string SECTION_PALIER = "  - Nettoyage des portes d'ascenseur de chaque palier\r\n  - Nettoyage des rails d'ascenseur de chaque palier";
    }
}
