using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using TemplateEngine.Docx;

namespace DevisGenerator.Pages
{
    /// <summary>
    /// Logique d'interaction pour FormulaireDevisCoproType.xaml
    /// </summary>
    public partial class FormulaireDevisCoproType : Window
    {
        private static readonly Regex DecimalInputRegex = new(@"^\d*(,\d*)?$");

        private sealed class LigneTarif : INotifyPropertyChanged
        {
            private string _libelle = string.Empty;
            private string _prixHT = string.Empty;
            private decimal _tauxTva = 0.20m;

            public event PropertyChangedEventHandler? PropertyChanged;

            public string Libelle
            {
                get => _libelle;
                set
                {
                    if (_libelle == value)
                        return;
                    _libelle = value;
                    OnPropertyChanged();
                }
            }

            public string PrixHT
            {
                get => _prixHT;
                set
                {
                    if (_prixHT == value)
                        return;
                    _prixHT = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(MontantHT));
                    OnPropertyChanged(nameof(MontantTVA));
                    OnPropertyChanged(nameof(MontantTTC));
                }
            }

            /// <summary>Taux de TVA appliqué à cette ligne, exprimé en fraction (0,20 = 20%).</summary>
            public decimal TauxTva
            {
                get => _tauxTva;
                set
                {
                    if (_tauxTva == value)
                        return;
                    _tauxTva = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(MontantTVA));
                    OnPropertyChanged(nameof(MontantTTC));
                }
            }

            public decimal MontantHT => ParseDecimal(PrixHT);
            public decimal MontantTVA => MontantHT * TauxTva;
            public decimal MontantTTC => MontantHT + MontantTVA;

            internal static decimal ParseDecimal(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return 0m;

                var normalized = value.Replace(" ", string.Empty);
                if (decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.CurrentCulture, out var result))
                    return result;
                if (decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.GetCultureInfo("fr-FR"), out result))
                    return result;
                if (decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out result))
                    return result;

                return 0m;
            }

            private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private readonly ObservableCollection<LigneTarif> _lignesTarif = new();

        public FormulaireDevisCoproType()
        {
            InitializeComponent();
            InitDevis();
            InitialiserGrilleTarifaire();
        }

        private BitmapSource? PhotoAscenceur;
        private BitmapSource? PhotoHall;
        private BitmapSource? PhotoCageEscaliers;
        private BitmapSource? PhotoPaliers;
        private BitmapSource? PhotoGarage;
        private BitmapSource? PhotoAbords;
        private BitmapSource? PhotoLocalConteneurVideOrdures;

        public void InitDevis()
        {
            _ = new DevisGenerator.Models.Devis();
        }

        private void HandlePhotoButtonClick(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn)
                return;

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files|*.*",
                Title = "Sélectionnez une image"
            };

            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var path = dlg.FileName;
                btn.ToolTip = path;
                btn.Content = Path.GetFileName(path);

                BitmapImage bmp = new BitmapImage();
                using (var fs = File.OpenRead(path))
                {
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = fs;
                    bmp.EndInit();
                    bmp.Freeze();
                }

                const int maxWidthPx = 140;
                const int maxHeightPx = 220;
                BitmapSource finalImage = bmp;

                if (bmp.PixelWidth > 0 && bmp.PixelHeight > 0)
                {
                    double scaleX = (double)maxWidthPx / bmp.PixelWidth;
                    double scaleY = (double)maxHeightPx / bmp.PixelHeight;
                    double scale = Math.Min(Math.Min(scaleX, scaleY), 1.0);
                    if (scale < 1.0)
                    {
                        var transform = new ScaleTransform(scale, scale);
                        var tb = new TransformedBitmap(bmp, transform);
                        tb.Freeze();
                        finalImage = tb;
                    }
                }

                switch (btn.Name)
                {
                    case "btnPhotoAscenceur":
                        PhotoAscenceur = finalImage;
                        break;
                    case "btnPhotoHall":
                        PhotoHall = finalImage;
                        break;
                    case "btnPhotoCageEscaliers":
                        PhotoCageEscaliers = finalImage;
                        break;
                    case "btnPhotoPaliers":
                        PhotoPaliers = finalImage;
                        break;
                    case "btnPhotoGarage":
                        PhotoGarage = finalImage;
                        break;
                    case "btnPhotoAbords":
                        PhotoAbords = finalImage;
                        break;
                    case "btnPhotoLocalConteneurVideOrdures":
                        PhotoLocalConteneurVideOrdures = finalImage;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Impossible d'ouvrir l'image : " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnPhotoAscenceur_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoHall_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoCageEscaliers_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoPaliers_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoGarage_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoAbords_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);
        private void BtnPhotoPhotoLocalConteneurVideOrdures_Click(object sender, RoutedEventArgs e) => HandlePhotoButtonClick(sender, e);

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValiderGrilleTarifaire())
                    return;

                string appDir = AppContext.BaseDirectory;
                string folderTemplate = "Templates";
                string templateFileName = "copro_devis_type_V2.docx";
                string templatePath = Path.Combine(appDir, folderTemplate, templateFileName);

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show(this, $"Modèle introuvable : {templatePath}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var saveDlg = new Microsoft.Win32.SaveFileDialog()
                {
                    Title = "Enregistrer le devis modifié",
                    Filter = "Documents Word|*.docx",
                    FileName = Path.GetFileNameWithoutExtension(templateFileName) + DateTime.Now.ToString("dd_MM_yyyy HH_mm_ss") + ".docx",
                    OverwritePrompt = true
                };

                if (saveDlg.ShowDialog() != true)
                    return;

                string outputPath = saveDlg.FileName;
                File.Copy(templatePath, outputPath, overwrite: true);

                GenererDevis(outputPath);

                MessageBox.Show(this, $"Document créé et modifié :{Environment.NewLine}{outputPath}{Environment.NewLine}{Environment.NewLine}Le document va être ouvert. Fermez-le après vos modifications pour poursuivre.", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);

                OuvrirDocument(outputPath);
                await AttendreFermetureDocumentAsync(outputPath);
                ProposerEnvoiParMail(outputPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erreur lors de l'insertion du texte : " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Remplit la copie du modèle à partir des valeurs du formulaire : retire d'abord les
        /// sections/photos optionnelles non renseignées (paragraphe entier supprimé, sans ligne
        /// vide résiduelle), redimensionne les photos choisies, puis injecte champs, images et
        /// tableau tarifaire via TemplateEngine.Docx (content controls Word).
        /// </summary>
        /// <param name="outputPath">Chemin de la copie du modèle .docx à remplir.</param>
        private void GenererDevis(string outputPath)
        {
            // Champs du modèle sans saisie correspondante dans le formulaire (nom/adresse du
            // syndic) : toujours vides, donc toujours retirés entièrement.
            OpenXmlWordHelpers.SupprimerParagrapheParTag(outputPath, "nomSyndic");
            OpenXmlWordHelpers.SupprimerParagrapheParTag(outputPath, "adresseSyndic");

            // Sections/lignes de checklist qui doivent disparaître entièrement si la case
            // correspondante n'est pas cochée (leur paragraphe est autonome dans le modèle).
            SupprimerSiNonCoche(outputPath, "ShowTitreAscenceur", chkShowAscenceur);
            SupprimerSiNonCoche(outputPath, "ShowAscenceur", chkShowAscenceur);
            SupprimerSiNonCoche(outputPath, "frenquenceAscenceur", chkShowAscenceur);
            SupprimerSiNonCoche(outputPath, "ShowTitreCaves", chkShowCavesAcces);
            SupprimerSiNonCoche(outputPath, "ShowCaves", chkShowCavesAcces);
            SupprimerSiNonCoche(outputPath, "frenquenceCavesEtAccesCaves", chkShowCavesAcces);
            SupprimerSiNonCoche(outputPath, "ShowPalier", chkShowAccesGarage);
            SupprimerSiNonCoche(outputPath, "NettoyagePortesAscenceur", nettoyagePortesAscenceurCheckbox);
            SupprimerSiNonCoche(outputPath, "NettoyageRailsAscenceur", nettoyageRailsAscenceurCheckbox);
            SupprimerSiNonCoche(outputPath, "ShowTitreAccesGarage", chkShowAccesGarage);
            SupprimerSiNonCoche(outputPath, "ShowAccesGarage", chkShowAccesGarage);
            SupprimerSiNonCoche(outputPath, "ShowTitreGarage", chkShowGarage);
            SupprimerSiNonCoche(outputPath, "ShowGarage", chkShowGarage);
            SupprimerSiNonCoche(outputPath, "frenquenceAccesGarage", chkShowAccesGarage);
            SupprimerSiNonCoche(outputPath, "frenquenceGarage", chkShowGarage);
            SupprimerSiNonCoche(outputPath, "EnlevementDetritutsFeuillesVoieAccesGarage", enlevementDetritutsFeuillesVoieAccesGarageCheckbox);
            SupprimerSiNonCoche(outputPath, "ShowTitreAbordsAcces", chkShowAbords);
            SupprimerSiNonCoche(outputPath, "ShowAbordsAcces", chkShowAbords);
            SupprimerSiNonCoche(outputPath, "frenquenceAbordsEtAcces", chkShowAbords);
            SupprimerSiNonCoche(outputPath, "ShowTitreLocalConteneurs", chkShowLocalConteneurs);
            SupprimerSiNonCoche(outputPath, "ShowLocalConteneurs", chkShowLocalConteneurs);
            SupprimerSiNonCoche(outputPath, "frenquenceLocalConteneurVideOrdures", chkShowLocalConteneurs);
            SupprimerSiNonCoche(outputPath, "SortieConteneurs", sortieConteneurCheckbox);
            SupprimerSiNonCoche(outputPath, "SortieSacs", sortieSacsCheckbox);
            SupprimerSiNonCoche(outputPath, "TransportEau", chkTransportEau);

            // Images fixes conditionnelles : déjà intégrées au modèle, on ne retire que si absentes.
            SupprimerSiNonCoche(outputPath, "ShowPrestationComplementaires", chkShowPrestationComplementaires);
            SupprimerSiNonCoche(outputPath, "CoproTravaux", coproTravauxCheckbox);
            SupprimerSiNonCoche(outputPath, "ActionNuisibles", actionNuisiblesCheckbox);
            SupprimerSiNonCoche(outputPath, "Relamping", relampingCheckbox);
            SupprimerSiNonCoche(outputPath, "Tracabilite", tracabiliteCheckbox);
            SupprimerSiNonCoche(outputPath, "ControleQualite", controleQualiteCheckbox);

            // Photos choisies par l'utilisateur : redimensionnées si présentes, retirées sinon.
            var photos = new (string Tag, BitmapSource? Image)[]
            {
                ("PhotoAscenceur", PhotoAscenceur),
                ("PhotoHall", PhotoHall),
                ("PhotoCageEscaliers", PhotoCageEscaliers),
                ("PhotoPaliers", PhotoPaliers),
                ("PhotoGarage", PhotoGarage),
                ("AbordsEtAcces", PhotoAbords),
                ("PhotoLocal", PhotoLocalConteneurVideOrdures),
            };

            foreach (var (tag, image) in photos)
            {
                if (image != null)
                    OpenXmlWordHelpers.RemplacerImagePlaceholder(outputPath, tag, OpenXmlWordHelpers.EncoderImagePng(image), image.PixelWidth, image.PixelHeight);
                else
                    OpenXmlWordHelpers.SupprimerParagrapheParTag(outputPath, tag);
            }

            string frenquenceHallEntreeFinale = GetComboOrManualValue(frenquenceHallEntree, frenquenceHallEntree_Manual, "Autre");
            string frenquenceAscenceurFinale = GetComboOrManualValue(frenquenceAscenceur, frenquenceAscenceur_Manual, "Autre");
            string frenquenceCageDEscaliersFinale = GetComboOrManualValue(frenquenceCageDEscaliers, frenquenceCageDEscaliers_Manual, "Autre");
            string frenquencePaliersFinale = GetComboOrManualValue(frenquencePaliers, frenquencePaliers_Manual, "Autre");
            string frenquenceCavesEtAccesCavesFinale = GetComboOrManualValue(frenquenceCavesEtAccesCaves, frenquenceCavesEtAccesCaves_Manual, "Autre");
            string frenquenceAccesGarageFinale = GetComboOrManualValue(frenquenceAccesGarage, frenquenceAccesGarage_Manual, "Autre");
            string frenquenceGarageFinale = GetComboOrManualValue(frenquenceGarage, frenquenceGarage_Manual, "Autre");
            string frenquenceAbordsEtAccesFinale = GetComboOrManualValue(frenquenceAbordsEtAcces, frenquenceAbordsEtAcces_Manual, "Autre");
            string frenquenceLocalConteneurVideOrduresFinale = GetComboOrManualValue(frenquenceLocalConteneurVideOrdures, frenquenceLocalConteneurVideOrdures_Manual, "Autre");
            string frenquenceOrduresMenageresFinale = GetComboOrManualValue(frenquenceOrduresMenageres, frenquenceOrduresMenageres_Manual, "Autre");

            var elements = new List<IContentItem>
            {
                Champ("NettoyageMarches", GetCheckboxValue(nettoyageMarchesCheckbox, Constantes.Constantes.NETTOYAGE_MARCHES)),
                Champ("NettoyageCuivres", GetCheckboxValue(nettoyageCuivreCheckbox, Constantes.Constantes.NETTOYAGE_CUIVRES)),
                Champ("NettoyagePorteTambour", GetCheckboxValue(nettoyagePorteTambourCheckbox, Constantes.Constantes.NETTOYAGE_PORTE_TAMBOUR)),
                Champ("RentreeConteneurs", GetCheckboxValue(rentreeConteneursCheckbox, Constantes.Constantes.RENTREE_CONTENEURS)),

                Champ("frenquenceHallEntree", frenquenceHallEntreeFinale),
                Champ("frenquenceCageDEscaliers", frenquenceCageDEscaliersFinale),
                Champ("frenquencePaliers", frenquencePaliersFinale),
                Champ("frenquenceOrduresMenageres", frenquenceOrduresMenageresFinale),

                Champ("clientNom", clientNom?.Text?.Trim() ?? string.Empty),
                Champ("clientAdresse", clientAdresse?.Text?.Trim() ?? string.Empty),
                Champ("numeroOpportunite", numeroOpportunite?.Text?.Trim() ?? string.Empty),
                Champ("nomOpportunite", nomOpportunite?.Text?.Trim() ?? string.Empty),
                Champ("referenceDevisClient", referenceDevisClient?.Text?.Trim() ?? string.Empty),
                Champ("dateDevis", DateTime.Now.ToString("d")),

                new TableContent("GrilleTarif", CreerLignesDocumentTarif().Select(l =>
                    new TableRowContent(new IContentItem[]
                    {
                        new FieldContent("Libelle", l.Libelle),
                        new FieldContent("PrixHT", l.PrixHT),
                        new FieldContent("TVA", l.TVA),
                        new FieldContent("PrixTTC", l.TTC),
                    }))),
                Champ("TotalHT", CalculerTotalHT().ToString("N2", CultureInfo.CurrentCulture) + " €"),
                Champ("TotalTVA", CalculerTotalTVA().ToString("N2", CultureInfo.CurrentCulture) + " €"),
                Champ("TotalTTC", CalculerTotalTTC().ToString("N2", CultureInfo.CurrentCulture) + " €"),
                Champ("TvaHeader", "TVA (" + (ObtenirTauxTvaSaisi() * 100m).ToString("0.##", CultureInfo.CurrentCulture) + "%)"),
            };

            // Sections autonomes : ajoutées seulement si cochées (sinon déjà retirées ci-dessus).
            if (chkShowAscenceur?.IsChecked == true) elements.Add(Champ("ShowTitreAscenceur", Constantes.Constantes.TITRE_SECTION_NETTOYAGE_CABINE_ASCENCEUR));
            if (chkShowAscenceur?.IsChecked == true) elements.Add(Champ("ShowAscenceur", Constantes.Constantes.SECTION_NETTOYAGE_CABINE_ASCENCEUR));
            if (chkShowAscenceur?.IsChecked == true) elements.Add(Champ("frenquenceAscenceur", frenquenceAscenceurFinale));
            if (chkShowCavesAcces?.IsChecked == true) elements.Add(Champ("ShowTitreCaves", Constantes.Constantes.TITRE_SECTION_CAVES_ET_ACCES_CAVES));
            if (chkShowCavesAcces?.IsChecked == true) elements.Add(Champ("ShowCaves", Constantes.Constantes.SECTION_CAVES_ET_ACCES_CAVES));
            if (chkShowCavesAcces?.IsChecked == true) elements.Add(Champ("frenquenceCavesEtAccesCaves", frenquenceCavesEtAccesCavesFinale));
            if (chkShowAccesGarage?.IsChecked == true) elements.Add(Champ("ShowPalier", Constantes.Constantes.SECTION_PALIER));
            if (nettoyagePortesAscenceurCheckbox?.IsChecked == true) elements.Add(Champ("NettoyagePortesAscenceur", Constantes.Constantes.NETTOYAGE_PORTES_ASCENCEUR));
            if (nettoyageRailsAscenceurCheckbox?.IsChecked == true) elements.Add(Champ("NettoyageRailsAscenceur", Constantes.Constantes.NETTOYAGE_RAILS_ASCENCEUR));
            if (chkShowAccesGarage?.IsChecked == true) elements.Add(Champ("ShowTitreAccesGarage", Constantes.Constantes.TITRE_SECTION_ACCES_GARAGE));
            if (chkShowAccesGarage?.IsChecked == true) elements.Add(Champ("ShowAccesGarage", Constantes.Constantes.SECTION_ACCES_GARAGE));
            if (chkShowGarage?.IsChecked == true) elements.Add(Champ("ShowTitreGarage", Constantes.Constantes.TITRE_SECTION_GARAGE));
            if (chkShowGarage?.IsChecked == true) elements.Add(Champ("ShowGarage", Constantes.Constantes.SECTION_GARAGE));
            if (chkShowAccesGarage?.IsChecked == true) elements.Add(Champ("frenquenceAccesGarage", frenquenceAccesGarageFinale));
            if (chkShowGarage?.IsChecked == true) elements.Add(Champ("frenquenceGarage", frenquenceGarageFinale));
            if (enlevementDetritutsFeuillesVoieAccesGarageCheckbox?.IsChecked == true) elements.Add(Champ("EnlevementDetritutsFeuillesVoieAccesGarage", Constantes.Constantes.ENLEVEMENT_DETRITUTS_FEUILLES_VOIE_ACCES_GARAGE));
            if (chkShowAbords?.IsChecked == true) elements.Add(Champ("ShowTitreAbordsAcces", Constantes.Constantes.TITRE_SECTION_ABORDS_ACCES));
            if (chkShowAbords?.IsChecked == true) elements.Add(Champ("ShowAbordsAcces", Constantes.Constantes.SECTION_ABORDS_ACCES));
            if (chkShowAbords?.IsChecked == true) elements.Add(Champ("frenquenceAbordsEtAcces", frenquenceAbordsEtAccesFinale));
            if (chkShowLocalConteneurs?.IsChecked == true) elements.Add(Champ("ShowTitreLocalConteneurs", Constantes.Constantes.TITRE_SECTION_LOCAL_CONTENEUR_VIDE_ORDURE));
            if (chkShowLocalConteneurs?.IsChecked == true) elements.Add(Champ("ShowLocalConteneurs", Constantes.Constantes.SECTION_LOCAL_CONTENEUR_VIDE_ORDURE));
            if (chkShowLocalConteneurs?.IsChecked == true) elements.Add(Champ("frenquenceLocalConteneurVideOrdures", frenquenceLocalConteneurVideOrduresFinale));
            if (sortieConteneurCheckbox?.IsChecked == true) elements.Add(Champ("SortieConteneurs", Constantes.Constantes.SORTIE_CONTENEUR));
            if (sortieSacsCheckbox?.IsChecked == true) elements.Add(Champ("SortieSacs", Constantes.Constantes.SORTIE_SACS));

            // Les photos ont déjà été injectées directement via OpenXmlWordHelpers.RemplacerImagePlaceholder
            // ci-dessus (TemplateEngine.Docx place les nouvelles images au mauvais endroit dans le paquet
            // .docx via ImageContent, ce qui rend le document illisible par Word).

            using (var processor = new TemplateProcessor(outputPath).SetRemoveContentControls(true).SetNoticeAboutErrors(false))
            {
                processor.FillContent(new Content(elements.ToArray()));
                processor.SaveChanges();
            }

            OpenXmlWordHelpers.ReparerRetoursALaLigne(outputPath);
        }

        /// <summary>
        /// Construit un champ texte pour TemplateEngine.Docx en neutralisant d'abord ses éventuels
        /// retours à la ligne (contournement du bug de la librairie sur les valeurs multi-lignes).
        /// </summary>
        /// <param name="tag">Tag du content control Word correspondant.</param>
        /// <param name="valeur">Valeur brute à injecter.</param>
        /// <returns>Le <see cref="FieldContent"/> prêt à être ajouté au contenu du document.</returns>
        private static FieldContent Champ(string tag, string? valeur)
        {
            return new FieldContent(tag, OpenXmlWordHelpers.SecuriserValeurMultiligne(valeur));
        }

        /// <summary>
        /// Retire du document le paragraphe portant le tag donné si la case à cocher associée
        /// n'est pas cochée (sections et images optionnelles).
        /// </summary>
        /// <param name="outputPath">Chemin du document .docx en cours de préparation.</param>
        /// <param name="tag">Tag du content control à retirer si non coché.</param>
        /// <param name="checkbox">Case à cocher pilotant la présence de ce contenu.</param>
        private static void SupprimerSiNonCoche(string outputPath, string tag, CheckBox? checkbox)
        {
            if (checkbox?.IsChecked != true)
                OpenXmlWordHelpers.SupprimerParagrapheParTag(outputPath, tag);
        }

        private static void OuvrirDocument(string outputPath)
        {
            Process.Start(new ProcessStartInfo(outputPath)
            {
                UseShellExecute = true
            });
        }

        private static async Task AttendreFermetureDocumentAsync(string outputPath)
        {
            string dossier = Path.GetDirectoryName(outputPath) ?? string.Empty;
            string nomFichier = Path.GetFileName(outputPath);
            string fichierVerrou = Path.Combine(dossier, "~$" + nomFichier);
            bool documentOuvert = false;
            DateTime limiteDetection = DateTime.UtcNow.AddSeconds(15);

            while (DateTime.UtcNow < limiteDetection)
            {
                if (File.Exists(fichierVerrou) || !PeutOuvrirFichier(outputPath))
                {
                    documentOuvert = true;
                    break;
                }

                await Task.Delay(500);
            }

            if (!documentOuvert)
                return;

            while (File.Exists(fichierVerrou) || !PeutOuvrirFichier(outputPath))
            {
                await Task.Delay(1000);
            }
        }

        private void ProposerEnvoiParMail(string outputPath)
        {
            string emailClient = emailClientTextBox?.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(emailClient))
            {
                emailClient = DemanderEmailClient();
                if (string.IsNullOrWhiteSpace(emailClient))
                    return;
            }

            if (!EmailEstValide(emailClient))
            {
                MessageBox.Show(this, "Le document a été fermé, mais l'e-mail client saisi n'est pas valide.", "Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show(this, $"Le document a bien été fermé.{Environment.NewLine}{Environment.NewLine}Souhaitez-vous préparer un e-mail pour {emailClient} ?", "Envoi par mail", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            (string? cheminPdf, string? erreurConversion) = ConvertirDocxEnPdf(outputPath);
            string documentAEnvoyer = cheminPdf ?? outputPath;

            if (cheminPdf == null)
            {
                MessageBox.Show(this, $"Impossible de convertir le document en PDF :{Environment.NewLine}{erreurConversion}{Environment.NewLine}{Environment.NewLine}Le document Word sera joint à la place.", "Information", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            if (TenterOuvertureOutlook(emailClient, documentAEnvoyer))
                return;

            string sujet = Uri.EscapeDataString("Devis copropriété");
            string corps = Uri.EscapeDataString($"Bonjour,{Environment.NewLine}{Environment.NewLine}Veuillez trouver le devis en pièce jointe.{Environment.NewLine}{Environment.NewLine}Document : {documentAEnvoyer}");
            Process.Start(new ProcessStartInfo($"mailto:{emailClient}?subject={sujet}&body={corps}")
            {
                UseShellExecute = true
            });

            MessageBox.Show(this, $"Le client mail a été ouvert sans pièce jointe automatique.{Environment.NewLine}{Environment.NewLine}Ajoutez manuellement ce document :{Environment.NewLine}{documentAEnvoyer}", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// Convertit un document Word en PDF via l'automation COM de Microsoft Word (même
        /// mécanisme de late-binding que l'envoi Outlook, sans dépendance de compilation).
        /// </summary>
        /// <param name="cheminDocx">Chemin du document .docx à convertir.</param>
        /// <returns>Le chemin du PDF généré à côté du document source (et null en erreur), ou le chemin null accompagné du message d'erreur si la conversion a échoué.</returns>
        private static (string? CheminPdf, string? Erreur) ConvertirDocxEnPdf(string cheminDocx)
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                return (null, "Microsoft Word n'est pas installé ou n'est pas enregistré comme application COM sur ce poste.");

            dynamic? word = null;
            dynamic? document = null;
            try
            {
                string cheminPdf = Path.ChangeExtension(cheminDocx, ".pdf");

                word = Activator.CreateInstance(wordType)!;
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone

                document = word.Documents.Open(cheminDocx, ReadOnly: false, Visible: false);
                document.ExportAsFixedFormat(cheminPdf, 17); // wdExportFormatPDF

                return File.Exists(cheminPdf) ? (cheminPdf, null) : (null, "Le fichier PDF n'a pas été créé (aucune exception levée, mais le fichier est absent).");
            }
            catch (Exception ex)
            {
                return (null, $"{ex.GetType().Name} : {ex.Message}");
            }
            finally
            {
                if (document != null)
                {
                    try { document.Close(false); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
                }
                if (word != null)
                {
                    try { word.Quit(false); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);
                }
            }
        }

        private string? DemanderEmailClient()
        {
            var dialog = new Window
            {
                Title = "E-mail client manquant",
                Width = 420,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false
            };

            var grid = new Grid { Margin = new System.Windows.Thickness(16) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new System.Windows.Controls.Label
            {
                Content = "Aucun e-mail client n'a été saisi. Veuillez le saisir :",
                Margin = new System.Windows.Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(label, 0);

            var textBox = new TextBox
            {
                Height = 26,
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new System.Windows.Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(textBox, 1);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(buttons, 2);

            string? result = null;

            var btnOk = new Button { Content = "Confirmer", MinWidth = 80, Margin = new System.Windows.Thickness(0, 0, 8, 0) };
            btnOk.Click += (_, _) =>
            {
                result = textBox.Text.Trim();
                dialog.DialogResult = true;
            };

            var btnCancel = new Button { Content = "Annuler", MinWidth = 80 };
            btnCancel.Click += (_, _) => { dialog.DialogResult = false; };

            buttons.Children.Add(btnOk);
            buttons.Children.Add(btnCancel);

            grid.Children.Add(label);
            grid.Children.Add(textBox);
            grid.Children.Add(buttons);

            dialog.Content = grid;
            textBox.Focus();

            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(result))
                return null;

            if (emailClientTextBox != null)
                emailClientTextBox.Text = result;

            return result;
        }

        private static bool TenterOuvertureOutlook(string emailClient, string outputPath)
        {
            try
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return false;

                dynamic outlook = Activator.CreateInstance(outlookType)!;
                dynamic mail = outlook.CreateItem(0); // olMailItem = 0

                mail.To = emailClient;
                mail.Subject = "Devis copropriété";
                mail.Body = $"Bonjour,{Environment.NewLine}{Environment.NewLine}Veuillez trouver le devis en pièce jointe.";
                mail.Attachments.Add(outputPath);
                mail.Display(false);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool EmailEstValide(string email)
        {
            try
            {
                _ = new MailAddress(email);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool PeutOuvrirFichier(string outputPath)
        {
            try
            {
                using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return true;
            }
            catch (IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }

        private static string GetCheckboxValue(CheckBox? checkbox, string value)
        {
            return checkbox?.IsChecked == true ? value : string.Empty;
        }

        private static string GetComboOrManualValue(ComboBox combo, TextBox manual, string marker = "Autre")
        {
            if (combo == null)
                return manual?.Text ?? string.Empty;

            var selected = combo.Text ?? string.Empty;
            if (!string.IsNullOrEmpty(selected) && selected.IndexOf(marker, StringComparison.CurrentCultureIgnoreCase) >= 0)
            {
                if (!string.IsNullOrEmpty(manual?.Text))
                    return manual.Text;
            }

            return selected;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void InitialiserGrilleTarifaire()
        {
            TarifItemsControl.ItemsSource = _lignesTarif;
            AjouterLigneTarif();
        }

        private void AjouterLigneTarif()
        {
            var ligne = new LigneTarif { TauxTva = ObtenirTauxTvaSaisi() };
            ligne.PropertyChanged += LigneTarif_PropertyChanged;
            _lignesTarif.Add(ligne);
            MettreAJourTotalTarif();
        }

        /// <summary>
        /// Lit le taux de TVA saisi par l'utilisateur (en pourcentage, ex. "20"), avec 20% par
        /// défaut si le champ est vide ou invalide.
        /// </summary>
        /// <returns>Le taux de TVA sous forme de fraction (0,20 pour 20%).</returns>
        private decimal ObtenirTauxTvaSaisi()
        {
            var pourcentage = LigneTarif.ParseDecimal(TauxTvaTextBox?.Text);
            return pourcentage > 0 ? pourcentage / 100m : 0.20m;
        }

        /// <summary>
        /// Répercute le taux de TVA saisi sur toutes les lignes existantes de la grille tarifaire
        /// et met à jour le total affiché.
        /// </summary>
        private void TauxTvaTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            decimal taux = ObtenirTauxTvaSaisi();
            foreach (var ligne in _lignesTarif)
                ligne.TauxTva = taux;

            MettreAJourTotalTarif();
        }

        private void RetirerLigneTarif(LigneTarif? ligne)
        {
            if (ligne == null)
                return;

            if (_lignesTarif.Count <= 1)
            {
                MessageBox.Show(this, "Au moins une ligne tarifaire doit être saisie.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ligne.PropertyChanged -= LigneTarif_PropertyChanged;
            _lignesTarif.Remove(ligne);
            MettreAJourTotalTarif();
        }

        private void AjouterLigneTarif_Click(object sender, RoutedEventArgs e)
        {
            AjouterLigneTarif();
        }

        private void SupprimerLigneTarif_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: LigneTarif ligne })
            {
                RetirerLigneTarif(ligne);
            }
        }

        private void LigneTarif_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            MettreAJourTotalTarif();
        }

        private void MettreAJourTotalTarif()
        {
            if (TotalTarifTextBlock != null)
            {
                TotalTarifTextBlock.Text = CalculerTotalTTC().ToString("N2", CultureInfo.CurrentCulture) + " €";
            }
        }

        private decimal CalculerTotalHT() => _lignesTarif.Sum(l => l.MontantHT);
        private decimal CalculerTotalTVA() => _lignesTarif.Sum(l => l.MontantTVA);
        private decimal CalculerTotalTTC() => _lignesTarif.Sum(l => l.MontantTTC);

        private bool ValiderGrilleTarifaire()
        {
            if (_lignesTarif.Count == 0)
            {
                MessageBox.Show(this, "Au moins une ligne tarifaire doit être saisie.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (_lignesTarif.All(l => string.IsNullOrWhiteSpace(l.Libelle) && string.IsNullOrWhiteSpace(l.PrixHT)))
            {
                MessageBox.Show(this, "Veuillez renseigner au moins une ligne tarifaire.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private List<(string Libelle, string PrixHT, string TVA, string TTC)> CreerLignesDocumentTarif()
        {
            return _lignesTarif
                .Where(l => !string.IsNullOrWhiteSpace(l.Libelle) || !string.IsNullOrWhiteSpace(l.PrixHT))
                .Select(l => (
                    l.Libelle ?? string.Empty,
                    l.MontantHT.ToString("N2", CultureInfo.CurrentCulture) + " €",
                    l.MontantTVA.ToString("N2", CultureInfo.CurrentCulture) + " €",
                    l.MontantTTC.ToString("N2", CultureInfo.CurrentCulture) + " €"))
                .ToList();
        }

        private void DecimalTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (sender is not TextBox textBox)
                return;

            var proposedText = GetProposedText(textBox, e.Text);
            e.Handled = !IsValidDecimalInput(proposedText);
        }

        private void DecimalTextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (sender is not TextBox textBox)
                return;

            if (!e.SourceDataObject.GetDataPresent(DataFormats.Text))
            {
                e.CancelCommand();
                return;
            }

            var pastedText = e.SourceDataObject.GetData(DataFormats.Text) as string ?? string.Empty;
            var proposedText = GetProposedText(textBox, pastedText);

            if (!IsValidDecimalInput(proposedText))
            {
                e.CancelCommand();
            }
        }

        private static string GetProposedText(TextBox textBox, string input)
        {
            var currentText = textBox.Text ?? string.Empty;
            var selectionStart = textBox.SelectionStart;
            var selectionLength = textBox.SelectionLength;

            if (selectionLength > 0)
            {
                currentText = currentText.Remove(selectionStart, selectionLength);
            }

            return currentText.Insert(selectionStart, input);
        }

        private static bool IsValidDecimalInput(string text)
        {
            return DecimalInputRegex.IsMatch(text);
        }
    }
}
