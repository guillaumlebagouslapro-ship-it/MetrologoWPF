using System;
using System.Collections.Generic;
using System.Linq;
using Metrologo.Models;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Vérifie qu'une série de mesures tombe dans le domaine couvert par le module
    /// d'incertitude sélectionné. Centralise la conversion « mesure brute → fréquence réelle
    /// → valeur de lookup (Hz ou tr/min) » pour ne pas la dupliquer dans l'orchestrateur.
    /// La même conversion affine existe dans <c>ExcelService.ConvertirEnFreqReelle</c> et
    /// <c>ExcelInteropHost.ConvertirEnFreqReelleLocal</c> (chemins d'écriture des coeffs).
    /// </summary>
    public static class IncertitudeCouverture
    {
        /// <summary>Verdict de couverture pour une gate.</summary>
        public readonly struct Resultat
        {
            public CouvertureModule Couverture { get; init; }
            /// <summary>Valeur effectivement comparée au module (Hz, ou tr/min pour les tachys).</summary>
            public double ValeurLookup { get; init; }
            /// <summary>Unité de <see cref="ValeurLookup"/> : « Hz » ou « tr/min ».</summary>
            public string Unite { get; init; }

            /// <summary>Vrai si la valeur sort du domaine d'un module pourtant présent.</summary>
            public bool EstHorsPlage => Couverture == CouvertureModule.HorsPlage;
        }

        /// <summary>
        /// Évalue la moyenne des <paramref name="valeurs"/> face au module d'incertitude
        /// principal de la mesure. Si aucun module n'est sélectionné (ou pas de valeurs),
        /// renvoie <see cref="CouvertureModule.Couvert"/> (comportement historique : on laisse
        /// passer avec les coefficients par défaut).
        /// </summary>
        public static Resultat Verifier(Mesure mesure, IReadOnlyList<double> valeurs, double tempsGateSecondes)
        {
            if (mesure == null
                || string.IsNullOrEmpty(mesure.NumModuleIncertitude)
                || valeurs == null
                || valeurs.Count == 0)
            {
                return new Resultat { Couverture = CouvertureModule.Couvert };
            }

            double moyenneReelle = ConvertirEnFreqReelle(valeurs.Average(), mesure);

            // Les bornes des modules tachy/strobo sont saisies en tr/min côté admin :
            // on convertit Hz × 60 avant le lookup, comme dans les chemins d'écriture.
            bool uniteRpm = EnTetesMesureHelper.EstUniteRpm(mesure.TypeMesure);
            double valeurLookup = uniteRpm ? moyenneReelle * 60.0 : moyenneReelle;

            string fonction = IncertitudeFonctionHelper.NomFonction(mesure.TypeMesure);
            var couverture = ModulesIncertitudeService.VerifierCouverture(
                mesure.NumModuleIncertitude, mesure.TypeMesure, fonction, tempsGateSecondes, valeurLookup);

            return new Resultat
            {
                Couverture = couverture,
                ValeurLookup = valeurLookup,
                Unite = uniteRpm ? "tr/min" : "Hz"
            };
        }

        /// <summary>
        /// Reproduit la conversion de la colonne F (Fréq. Réelle) du template. Affine, donc
        /// commute avec la moyenne : <c>conv(AVERAGE(E)) == AVERAGE(conv(E))</c>.
        /// </summary>
        private static double ConvertirEnFreqReelle(double mesureBrute, Mesure mesure)
        {
            if (mesure.ModeMesure == ModeMesure.Direct) return mesureBrute;

            return (((mesureBrute - 10_000_000.0)
                    / (Math.Pow(10, mesure.IndexMultiplicateur) * 10_000_000.0)) + 1.0)
                   * mesure.FNominale;
        }
    }

    /// <summary>
    /// Levée lorsqu'une gate produit une valeur moyenne hors du domaine couvert par le module
    /// d'incertitude sélectionné. L'orchestrateur la rattrape pour supprimer la feuille de la
    /// mesure en cours (comme un arrêt) et remonter <see cref="MessageUtilisateur"/> au ViewModel.
    /// </summary>
    public sealed class MesureHorsModuleException : Exception
    {
        public string MessageUtilisateur { get; }

        public MesureHorsModuleException(string messageUtilisateur, string messageLog)
            : base(messageLog)
        {
            MessageUtilisateur = messageUtilisateur;
        }
    }
}
