using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace DevisGenerator.Converters
{
    /*
     PSEUDOCODE - Plan détaillé (en français) :
     1. Convert(value, targetType, parameter, culture)
        - But : détecter si la valeur liée représente une sélection de ComboBox contenant le texte "Autre".
        - Étapes :
          a) Si value est un ComboBoxItem et que ComboBoxItem.Content n'est pas null :
             - Récupérer le contenu en tant que chaîne s.
             - Vérifier si s contient "Autre" en ignorant la casse.
             - Si oui -> retourner Visibility.Visible.
          b) Sinon si value est déjà une string :
             - Vérifier si la chaîne contient "Autre" en ignorant la casse.
             - Si oui -> retourner Visibility.Visible.
          c) Sinon -> retourner Visibility.Collapsed.
     2. ConvertBack(value, targetType, parameter, culture)
        - But : la conversion inverse n'est pas pertinente ici (on ne peut pas déduire l'élément sélectionné à partir d'une Visibility).
        - Comportement : lancer une NotSupportedException.
    */

    public class ComboSelectionToVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Convertit la sélection d'un ComboBox en une valeur de visibilité.
        /// Renvoie <see cref="Visibility.Visible"/> si l'élément sélectionné contient le texte "Autre"
        /// (comparaison insensible à la casse), sinon <see cref="Visibility.Collapsed"/>.
        /// </summary>
        /// <param name="value">La valeur liée : généralement un <see cref="ComboBoxItem"/> ou une <see cref="string"/> représentant la sélection.</param>
        /// <param name="targetType">Type de la cible de la liaison (non utilisé).</param>
        /// <param name="parameter">Paramètre de conversion optionnel (non utilisé).</param>
        /// <param name="culture">Culture utilisée pour la conversion (non utilisée).</param>
        /// <returns><see cref="Visibility.Visible"/> si le texte contient "Autre", sinon <see cref="Visibility.Collapsed"/>.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // value is SelectedItem of ComboBox (usually a ComboBoxItem)
            if (value is ComboBoxItem cbi && cbi.Content != null)
            {
                var s = cbi.Content.ToString();
                if (s.IndexOf("Autre", StringComparison.OrdinalIgnoreCase) >= 0)
                    return Visibility.Visible;
            }
            // also if value is string
            if (value is string str)
            {
                if (str.IndexOf("Autre", StringComparison.OrdinalIgnoreCase) >= 0)
                    return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        /// <summary>
        /// Conversion inverse non prise en charge.
        /// Cette méthode lance toujours une <see cref="NotSupportedException"/> car il n'est pas possible
        /// de déduire l'élément de ComboBox à partir d'une valeur de visibilité.
        /// </summary>
        /// <param name="value">La valeur à convertir en sens inverse (non utilisée).</param>
        /// <param name="targetType">Type de la cible de la conversion inverse (non utilisé).</param>
        /// <param name="parameter">Paramètre de conversion optionnel (non utilisé).</param>
        /// <param name="culture">Culture utilisée pour la conversion (non utilisée).</param>
        /// <returns>Ne retourne jamais : lève une <see cref="NotSupportedException"/>.</returns>
        /// <exception cref="NotSupportedException">Toujours levée pour indiquer que la conversion inverse n'est pas supportée.</exception>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
