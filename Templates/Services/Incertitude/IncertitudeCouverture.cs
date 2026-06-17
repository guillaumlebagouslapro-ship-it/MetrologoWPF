using System;
using System.Collections.Generic;
using System.Linq;
using Metrologo.Models;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Vérifie qu'une série de mesures tombe dans le domaine couvert par le module
    /// d'incertitude sélectionné. Centralise la conversion « mesure brute → freq réelle
    /// → valeur de lookup (Hz ou tr/min) » pour éviter la duplication dans l'orchestrateur.
    /// Même conversion affine que <c>ExcelService.ConvertirEnFreqReelle</c> et
    /// <c>ExcelInteropHost.ConvertirEnFreqReelleLocal</c>.
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
        /// Évalue la moyenne des <paramref name="valeurs"/> face au module d'incertitude.
        /// Sans module sélectionné ou sans valeurs, retourne <see cref="CouvertureModule.Couvert"/>
        /// (comportement historique : passe avec coefficients par défaut).
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

            // Bornes tachy/strobo saisies en tr/min → Hz × 60 avant le lookup (comme côté écriture).
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

        /// <summary>Conversion colonne F (Fréq. Réelle) du template. Affine :
        /// <c>conv(AVERAGE(E)) == AVERAGE(conv(E))</c>.</summary>
        private static double ConvertirEnFreqReelle(double mesureBrute, Mesure mesure)
        {
            if (mesure.ModeMesure == ModeMesure.Direct) return mesureBrute;

            return (((mesureBrute - 10_000_000.0)
                    / (Math.Pow(10, mesure.IndexMultiplicateur) * 10_000_000.0)) + 1.0)
                   * mesure.FNominale;
        }
    }

    /// <summary>
    /// Levée quand la valeur moyenne d'une gate dépasse le domaine du module d'incertitude.
    /// L'orchestrateur la rattrape, supprime la feuille en cours et remonte
    /// <see cref="MessageUtilisateur"/> au ViewModel.
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
