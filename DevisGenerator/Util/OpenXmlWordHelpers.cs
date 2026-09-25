using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
/// Utilitaires OpenXml de préparation/réparation utilisés autour de TemplateEngine.Docx, qui
/// pilote l'essentiel de l'injection (champs, tableau tarifaire) via des content controls Word.
/// Ces méthodes gèrent ce que la librairie ne fait pas nativement, ou fait de façon incorrecte :
/// suppression complète d'un paragraphe optionnel (sans laisser de ligne vide), correction des
/// sauts de ligne dans les valeurs multi-lignes (contournement d'un bug produisant du XML invalide
/// quand une valeur de champ contient "\r\n"), et remplacement des images des photos (contournement
/// d'un bug qui place les nouvelles images au mauvais endroit dans le paquet .docx et rend le
/// document illisible par Word).
/// </summary>
public static class OpenXmlWordHelpers
{
    /// <summary>
    /// Caractère de remplacement temporaire pour les retours à la ligne à l'intérieur d'une
    /// valeur de champ, en attendant le passage de réparation <see cref="ReparerRetoursALaLigne"/>.
    /// </summary>
    private const string MarqueurRetourLigne = "@@BR@@";

    /// <summary>
    /// Neutralise les retours à la ligne d'une valeur de champ avant de la transmettre à
    /// TemplateEngine.Docx, qui génère du XML invalide (des &lt;w:t&gt; imbriqués) dès qu'une
    /// valeur contient "\r\n". Le marqueur inséré est reconverti par <see cref="ReparerRetoursALaLigne"/>
    /// après le remplissage du document.
    /// </summary>
    /// <param name="valeur">Valeur brute pouvant contenir des "\r\n" ou "\n".</param>
    /// <returns>Valeur avec les retours à la ligne remplacés par un marqueur sûr.</returns>
    public static string SecuriserValeurMultiligne(string? valeur)
    {
        if (string.IsNullOrEmpty(valeur))
            return string.Empty;

        return valeur.Replace("\r\n", MarqueurRetourLigne).Replace("\n", MarqueurRetourLigne);
    }

    /// <summary>
    /// Reconvertit les marqueurs posés par <see cref="SecuriserValeurMultiligne"/> en véritables
    /// sauts de ligne Word (&lt;w:br/&gt;), une fois que TemplateEngine.Docx a rempli le document.
    /// À appeler juste après <c>TemplateProcessor.SaveChanges()</c>, sur le fichier de sortie.
    /// </summary>
    /// <param name="cheminDocument">Chemin du document .docx déjà rempli à réparer.</param>
    public static void ReparerRetoursALaLigne(string cheminDocument)
    {
        using var doc = WordprocessingDocument.Open(cheminDocument, true);
        if (doc.MainDocumentPart?.Document.Body == null)
            return;

        foreach (var conteneur in EnumererConteneursTexte(doc))
        {
            foreach (var texte in conteneur.Descendants<WP.Text>().Where(t => t.Text.Contains(MarqueurRetourLigne)).ToList())
            {
                var run = texte.Ancestors<WP.Run>().First();
                var lignes = texte.Text.Split(MarqueurRetourLigne);
                var runProps = run.RunProperties?.CloneNode(true) as WP.RunProperties;

                var nouveauRun = new WP.Run();
                if (runProps != null)
                    nouveauRun.RunProperties = runProps;

                for (int i = 0; i < lignes.Length; i++)
                {
                    nouveauRun.AppendChild(new WP.Text(lignes[i]) { Space = SpaceProcessingModeValues.Preserve });
                    if (i < lignes.Length - 1)
                        nouveauRun.AppendChild(new WP.Break());
                }

                run.Parent!.ReplaceChild(nouveauRun, run);
            }
        }

        SauvegarderDocument(doc);
    }

    /// <summary>
    /// Supprime entièrement le paragraphe (ou la ligne de tableau) portant un content control
    /// dont le tag correspond, sans laisser de ligne vide résiduelle. Utilisé avant l'appel à
    /// TemplateProcessor pour les sections/photos optionnelles qui n'ont pas de valeur cette fois-ci.
    /// </summary>
    /// <param name="cheminDocument">Chemin du document .docx en cours de préparation.</param>
    /// <param name="tag">Tag du content control (paragraphe ou image) à retirer.</param>
    public static void SupprimerParagrapheParTag(string cheminDocument, string tag)
    {
        using var doc = WordprocessingDocument.Open(cheminDocument, true);
        if (doc.MainDocumentPart?.Document.Body == null || string.IsNullOrWhiteSpace(tag))
            return;

        foreach (var conteneur in EnumererConteneursTexte(doc))
        {
            OpenXmlElement? sdt =
                conteneur.Descendants<WP.SdtRun>().FirstOrDefault(s => s.SdtProperties?.GetFirstChild<WP.Tag>()?.Val == tag)
                ?? (OpenXmlElement?)conteneur.Descendants<WP.SdtBlock>().FirstOrDefault(s => s.SdtProperties?.GetFirstChild<WP.Tag>()?.Val == tag);

            if (sdt == null)
                continue;

            var parent = sdt.Parent;
            var paragraphe = sdt.Ancestors<WP.Paragraph>().FirstOrDefault();
            if (paragraphe != null)
            {
                parent = paragraphe.Parent;
                paragraphe.Remove();
            }
            else
            {
                sdt.Remove();
            }

            // Une cellule de tableau doit toujours contenir au moins un paragraphe (contrainte du
            // schéma Word) : la vider entièrement (cas des photos placées dans un tableau) rend le
            // document illisible par Word. On y remet un paragraphe vide si besoin.
            if (parent is WP.TableCell celluleTableau && !celluleTableau.Elements<WP.Paragraph>().Any())
            {
                celluleTableau.AppendChild(new WP.Paragraph());
            }
        }

        SauvegarderDocument(doc);
    }

    private static readonly XNamespace NsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace NsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>
    /// Remplace l'image d'un content control image par une nouvelle image (octets PNG) et fixe sa
    /// taille d'affichage (en pixels). Contourne volontairement le SDK OpenXml haut niveau
    /// (<c>MainDocumentPart.AddImagePart</c>, y compris via <c>ImageContent</c> de TemplateEngine.Docx) :
    /// sur cette version du runtime, l'ajout d'une image sans URI explicite la place systématiquement
    /// à la racine du paquet .docx (hors word/media/), ce qui rend le document illisible par Word.
    /// La méthode utilise directement l'API bas niveau <see cref="System.IO.Packaging.Package"/> avec
    /// un chemin explicite calculé pour ne jamais entrer en collision avec les images déjà présentes
    /// (y compris celles des en-têtes/pieds de page, invisibles pour le SDK haut niveau à ce stade).
    /// </summary>
    /// <param name="cheminDocument">Chemin du document .docx en cours de préparation.</param>
    /// <param name="tag">Tag du content control image dont l'image doit être remplacée.</param>
    /// <param name="octetsPng">Octets PNG de la nouvelle image.</param>
    /// <param name="largeurPx">Largeur finale à appliquer, en pixels.</param>
    /// <param name="hauteurPx">Hauteur finale à appliquer, en pixels.</param>
    public static void RemplacerImagePlaceholder(string cheminDocument, string tag, byte[] octetsPng, int largeurPx, int hauteurPx)
    {
        if (largeurPx <= 0 || hauteurPx <= 0)
            return;

        using var pkg = Package.Open(cheminDocument, FileMode.Open, FileAccess.ReadWrite);
        var docPartUri = new Uri("/word/document.xml", UriKind.Relative);
        var docPart = pkg.GetPart(docPartUri);

        XDocument xdoc;
        using (var s = docPart.GetStream(FileMode.Open, FileAccess.Read))
            xdoc = XDocument.Load(s);

        var sdt = xdoc.Descendants(NsW + "sdt")
            .FirstOrDefault(e => (string?)e.Element(NsW + "sdtPr")?.Element(NsW + "tag")?.Attribute(NsW + "val") == tag);
        if (sdt == null)
            return;

        var blip = sdt.Descendants(NsA + "blip").FirstOrDefault();
        var embedAttr = blip?.Attribute(NsR + "embed");
        if (embedAttr == null)
            return;

        string ancienId = embedAttr.Value;
        Uri? ancienneUri = null;
        if (docPart.RelationshipExists(ancienId))
        {
            ancienneUri = docPart.GetRelationship(ancienId).TargetUri;
            docPart.DeleteRelationship(ancienId);
        }

        string nomFichier = "image" + ProchainNumeroImageDisponible(pkg) + ".png";
        var nouvelleUri = new Uri("/word/media/" + nomFichier, UriKind.Relative);
        var nouvellePart = pkg.CreatePart(nouvelleUri, "image/png");
        using (var s = nouvellePart.GetStream(FileMode.Create, FileAccess.Write))
            s.Write(octetsPng, 0, octetsPng.Length);

        var nouvelleRelation = docPart.CreateRelationship(nouvelleUri, TargetMode.Internal, ImageRelationshipType);
        embedAttr.Value = nouvelleRelation.Id;

        if (ancienneUri != null && pkg.PartExists(ancienneUri))
            pkg.DeletePart(ancienneUri);

        long largeurEmu = PxToEmu(largeurPx);
        long hauteurEmu = PxToEmu(hauteurPx);

        foreach (var extent in sdt.Descendants(NsWp + "extent"))
        {
            extent.SetAttributeValue("cx", largeurEmu);
            extent.SetAttributeValue("cy", hauteurEmu);
        }
        foreach (var ext in sdt.Descendants(NsA + "ext"))
        {
            ext.SetAttributeValue("cx", largeurEmu);
            ext.SetAttributeValue("cy", hauteurEmu);
        }

        using (var s = docPart.GetStream(FileMode.Create, FileAccess.Write))
            xdoc.Save(s);

        pkg.Flush();
    }

    /// <summary>
    /// Calcule un numéro "imageN" garanti libre en examinant toutes les parties image déjà
    /// présentes dans le paquet (y compris celles des en-têtes/pieds de page), pour éviter toute
    /// collision de nom lors de l'ajout d'une nouvelle image via l'API bas niveau.
    /// </summary>
    /// <param name="pkg">Paquet .docx ouvert en lecture/écriture.</param>
    /// <returns>Le plus petit numéro d'image non encore utilisé dans le paquet.</returns>
    private static int ProchainNumeroImageDisponible(Package pkg)
    {
        int max = pkg.GetParts()
            .Select(p => Regex.Match(p.Uri.OriginalString, @"image(\d*)\.\w+$"))
            .Where(m => m.Success)
            .Select(m => m.Groups[1].Value.Length == 0 ? 1 : int.Parse(m.Groups[1].Value))
            .DefaultIfEmpty(0)
            .Max();

        return max + 1;
    }

    /// <summary>
    /// Encode une image bitmap en PNG en mémoire, pour alimenter un <c>ImageContent</c> de
    /// TemplateEngine.Docx (qui attend un tableau d'octets).
    /// </summary>
    /// <param name="image">Image source (déjà chargée/redimensionnée côté formulaire).</param>
    /// <returns>Octets PNG de l'image.</returns>
    public static byte[] EncoderImagePng(BitmapSource image)
    {
        using var ms = new MemoryStream();
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        encoder.Save(ms);
        return ms.ToArray();
    }


    private static IEnumerable<OpenXmlCompositeElement> EnumererConteneursTexte(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document.Body != null)
            yield return doc.MainDocumentPart.Document.Body;

        foreach (var headerPart in doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
        {
            if (headerPart.Header != null)
                yield return headerPart.Header;
        }

        foreach (var footerPart in doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
        {
            if (footerPart.Footer != null)
                yield return footerPart.Footer;
        }
    }

    private static void SauvegarderDocument(WordprocessingDocument doc)
    {
        doc.MainDocumentPart?.Document?.Save();

        foreach (var headerPart in doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
        {
            headerPart.Header?.Save();
        }

        foreach (var footerPart in doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
        {
            footerPart.Footer?.Save();
        }
    }

    private static long PxToEmu(int px) => px * 9525L;
}
