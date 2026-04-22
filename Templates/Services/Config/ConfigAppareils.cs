using System.Collections.Generic;
using Metrologo.Models;

namespace Metrologo.Services.Config
{
    /// <summary>
    /// Résultat du chargement de Metrologo.ini : trois fréquencemètres obligatoires
    /// + un multiplexeur optionnel (historiquement présent, plus utilisé aujourd'hui).
    /// </summary>
    public class ConfigAppareils
    {
        public required AppareilIEEE Stanford { get; init; }
        public required AppareilIEEE Racal { get; init; }
        public required AppareilIEEE Eip { get; init; }

        /// <summary>Multiplexeur HP59307A, optionnel. Null si absent de l'ini.</summary>
        public AppareilIEEE? Mux { get; init; }

        /// <summary>Avertissements non bloquants (ex. MUX absent, adresses en conflit).</summary>
        public required IReadOnlyList<string> Avertissements { get; init; }

        public AppareilIEEE Par(TypeAppareilIEEE type) => type switch
        {
            TypeAppareilIEEE.Stanford => Stanford,
            TypeAppareilIEEE.Racal    => Racal,
            TypeAppareilIEEE.EIP      => Eip,
            _ => throw new System.ArgumentOutOfRangeException(nameof(type))
        };
    }
}
