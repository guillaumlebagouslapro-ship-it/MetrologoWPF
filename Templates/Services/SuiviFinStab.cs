using System;
using System.Collections.Generic;
using System.Linq;

namespace Metrologo.Services
{
    /// <summary>
    /// Heure de fin estimée de la stabilité en cours, affichée dans le bandeau rouge « Mesure en
    /// cours » : « Fin estimée vers 14:35 · reste ≈ 1 h 12 », ou pour une série
    /// « Stab 2/3 · fin de la série vers 15:10 · reste ≈ 48 min ».
    /// <para>
    /// L'heure de fin est calculée au lancement de chaque stab puis RECALÉE à la fin de chaque gate
    /// (les durées réelles affinent l'estimation, voir <see cref="EstimationDureeStab"/>). Entre deux
    /// gates, le temps restant décompte simplement jusqu'à cette heure. Une fois l'heure dépassée,
    /// le bandeau dit « Fin imminente… » plutôt qu'un temps négatif.
    /// </para>
    /// Thread-safe : écrit par le thread de mesure, lu par le thread du bandeau.
    /// </summary>
    public static class SuiviFinStab
    {
        private static readonly object _verrou = new();
        private static bool _actif;
        private static List<int> _gates = new();
        private static int _nbMesures;
        private static int _gatesFaites;
        private static int _repetition = 1;
        private static int _nbRepetitions = 1;
        private static DateTime _finEstimee;

        static SuiviFinStab()
        {
            EstimationDureeStab.GateEnregistree += _ =>
            {
                lock (_verrou)
                {
                    if (!_actif) return;
                    _gatesFaites = Math.Min(_gatesFaites + 1, _gates.Count);
                    Recalculer();
                }
            };
        }

        /// <summary>Au lancement d'une stab (la <paramref name="repetition"/>-ième d'une série de
        /// <paramref name="nbRepetitions"/>, 1/1 pour une stab seule).</summary>
        public static void Demarrer(IEnumerable<int> gates, int nbMesures, int repetition, int nbRepetitions)
        {
            lock (_verrou)
            {
                _gates = gates?.ToList() ?? new List<int>();
                _nbMesures = nbMesures;
                _gatesFaites = 0;
                _repetition = Math.Max(1, repetition);
                _nbRepetitions = Math.Max(_repetition, nbRepetitions);
                _actif = _gates.Count > 0 && nbMesures > 0;
                if (_actif) Recalculer();
            }
        }

        public static void Arreter()
        {
            lock (_verrou) { _actif = false; }
        }

        /// <summary>Texte du bandeau, ou null s'il n'y a pas de stab en cours.</summary>
        public static string? TexteBandeau()
        {
            lock (_verrou)
            {
                if (!_actif) return null;

                string prefixe = _nbRepetitions > 1
                    ? $"Stab {_repetition}/{_nbRepetitions} · fin de la série"
                    : "Fin estimée";

                var reste = _finEstimee - DateTime.Now;
                if (reste <= TimeSpan.Zero)
                    return _nbRepetitions > 1
                        ? $"Stab {_repetition}/{_nbRepetitions} · fin imminente…"
                        : "Fin imminente…";

                string resteTxt = reste.TotalMinutes < 1 ? "moins d'1 min" : EstimationDureeStab.Formater(reste);
                return $"{prefixe} vers {EstimationDureeStab.HeureFin(reste)} · reste ≈ {resteTxt}";
            }
        }

        /// <summary>Gates restantes de la stab en cours + stabs restantes de la série. Sous verrou.</summary>
        private static void Recalculer()
        {
            var reste = EstimationDureeStab.EstimerBalayage(_gates.Skip(_gatesFaites), _nbMesures);
            int stabsSuivantes = _nbRepetitions - _repetition;
            if (stabsSuivantes > 0)
                reste += TimeSpan.FromTicks(EstimationDureeStab.EstimerBalayage(_gates, _nbMesures).Ticks * stabsSuivantes);
            _finEstimee = DateTime.Now + reste;
        }
    }
}
