using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Metrologo.Services;

namespace Metrologo.Behaviors
{
    /// <summary>
    /// Propriétés attachées restreignant un <see cref="TextBox"/> à la saisie numérique :
    /// frappes et collages non numériques rejetés avant d'atteindre le champ.
    /// Usage XAML : <c>behaviors:NumericInput.IsEnabled="True"</c>.
    /// La validation métier (bornes, signe) reste du ressort du ViewModel.
    /// </summary>
    public static class NumericInput
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled", typeof(bool), typeof(NumericInput),
                new PropertyMetadata(false, OnIsEnabledChanged));

        public static readonly DependencyProperty AllowDecimalProperty =
            DependencyProperty.RegisterAttached(
                "AllowDecimal", typeof(bool), typeof(NumericInput),
                new PropertyMetadata(true));

        public static readonly DependencyProperty AllowNegativeProperty =
            DependencyProperty.RegisterAttached(
                "AllowNegative", typeof(bool), typeof(NumericInput),
                new PropertyMetadata(true));

        public static bool GetIsEnabled(DependencyObject o) => (bool)o.GetValue(IsEnabledProperty);
        public static void SetIsEnabled(DependencyObject o, bool v) => o.SetValue(IsEnabledProperty, v);

        public static bool GetAllowDecimal(DependencyObject o) => (bool)o.GetValue(AllowDecimalProperty);
        public static void SetAllowDecimal(DependencyObject o, bool v) => o.SetValue(AllowDecimalProperty, v);

        public static bool GetAllowNegative(DependencyObject o) => (bool)o.GetValue(AllowNegativeProperty);
        public static void SetAllowNegative(DependencyObject o, bool v) => o.SetValue(AllowNegativeProperty, v);

        private static void OnIsEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBox tb) return;

            if ((bool)e.NewValue)
            {
                tb.PreviewTextInput += OnPreviewTextInput;
                DataObject.AddPastingHandler(tb, OnPaste);
                tb.PreviewKeyDown += OnPreviewKeyDown;
            }
            else
            {
                tb.PreviewTextInput -= OnPreviewTextInput;
                DataObject.RemovePastingHandler(tb, OnPaste);
                tb.PreviewKeyDown -= OnPreviewKeyDown;
            }
        }

        // Espace ne passe pas par PreviewTextInput : bloqué ici.
        private static void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
                e.Handled = true;
        }

        private static void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (sender is not TextBox tb) return;
            var prospective = BuildProspectiveText(tb, e.Text);
            if (!SaisieHelper.IsPlausibleNumericInput(prospective, GetAllowNegative(tb), GetAllowDecimal(tb)))
                e.Handled = true;
        }

        private static void OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (sender is not TextBox tb) return;

            if (!e.DataObject.GetDataPresent(DataFormats.Text))
            {
                e.CancelCommand();
                return;
            }

            var pasted = (string)e.DataObject.GetData(DataFormats.Text);
            var prospective = BuildProspectiveText(tb, pasted);
            if (!SaisieHelper.IsPlausibleNumericInput(prospective, GetAllowNegative(tb), GetAllowDecimal(tb)))
                e.CancelCommand();
        }

        /// <summary>Texte résultant si <paramref name="insert"/> remplaçait la sélection courante.</summary>
        private static string BuildProspectiveText(TextBox tb, string insert)
        {
            var text = tb.Text ?? string.Empty;
            int start = tb.SelectionStart;
            int len = tb.SelectionLength;
            return text.Substring(0, start) + insert + text.Substring(start + len);
        }
    }
}
