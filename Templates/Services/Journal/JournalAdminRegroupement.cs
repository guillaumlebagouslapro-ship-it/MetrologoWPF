using System.Text.RegularExpressions;

namespace Metrologo.Services.Journal
{
    public enum TypeBlocAdmin
    {
        /// <summary>Un admin s'est authentifié ; ses actions suivent.</summary>
        Connexion,
        /// <summary>Tentative(s) d'accès refusée(s).</summary>
        Refus,
        /// <summary>Changement fait par l'application elle-même (Besançon).</summary>
        Automatique,
        /// <summary>Action sans connexion admin retrouvée (anciennes entrées).</summary>
        SansConnexion,
    }

    /// <summary>Une action dans un bloc ; les répétitions consécutives sont fusionnées (×N).</summary>
    public sealed class ActionAdmin
    {
        public DateTime Heure { get; set; }
        public string Code { get; set; } = string.Empty;
        public string Libelle { get; set; } = string.Empty;
        public string Detail { get; set; } = string.Empty;
        public int Nombre { get; set; } = 1;

        public string HeureAffichee => Heure.ToString("HH:mm");
        public string LibelleAffiche => Nombre > 1 ? $"{Libelle} (×{Nombre})" : Libelle;
    }

    /// <summary>Un bloc du journal admin : qui, quand, sur quel poste, et ce qu'il a fait.</summary>
    public sealed class BlocJournalAdmin
    {
        public TypeBlocAdmin Type { get; set; }
        public DateTime Debut { get; set; }
        public string Utilisateur { get; set; } = string.Empty;
        public string Machine { get; set; } = string.Empty;
        public List<ActionAdmin> Actions { get; } = new();
        public int NbTentatives { get; set; }
        public string LoginsTentes { get; set; } = string.Empty;

        public bool AActions => Actions.Count > 0;

        public string Titre => Type switch
        {
            TypeBlocAdmin.Connexion     => $"Connexion de {Utilisateur}",
            TypeBlocAdmin.Refus         => NbTentatives > 1 ? $"Accès refusé ({NbTentatives} tentatives)" : "Accès refusé",
            TypeBlocAdmin.Automatique   => "Mise à jour automatique",
            _                           => string.IsNullOrWhiteSpace(Utilisateur) || Utilisateur == "?"
                                               ? "Action admin" : $"Action de {Utilisateur}",
        };

        public string SousTitre
        {
            get
            {
                string s = Debut.ToString("dd/MM/yyyy HH:mm");
                if (!string.IsNullOrWhiteSpace(Machine)) s += $" · {Machine}";
                if (Type == TypeBlocAdmin.Refus && LoginsTentes.Length > 0) s += $" · identifiant : {LoginsTentes}";
                return s;
            }
        }

        public string Resume => Type switch
        {
            TypeBlocAdmin.Refus => string.Empty,
            _ when Actions.Count == 0 => "Connexion seule",
            _ => $"{Actions.Sum(a => a.Nombre)} action{(Actions.Sum(a => a.Nombre) > 1 ? "s" : "")}",
        };

        public bool Contient(string q) =>
            Titre.Contains(q, StringComparison.OrdinalIgnoreCase)
            || SousTitre.Contains(q, StringComparison.OrdinalIgnoreCase)
            || Actions.Any(a => a.Libelle.Contains(q, StringComparison.OrdinalIgnoreCase)
                             || a.Detail.Contains(q, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Transforme les lignes brutes du journal d'audit en blocs lisibles : une connexion admin suivie
    /// de ses actions (ou « connexion seule »), les refus d'accès regroupés, les mises à jour
    /// automatiques à part. Les répétitions consécutives d'une même action sont fusionnées.
    /// </summary>
    public static class JournalAdminRegroupement
    {
        /// <summary>Au-delà, une action n'est plus rattachée à la dernière connexion du poste.</summary>
        private static readonly TimeSpan DureeMaxSession = TimeSpan.FromHours(12);
        /// <summary>Des refus espacés de moins que ça sont comptés dans le même bloc.</summary>
        private static readonly TimeSpan FenetreRefus = TimeSpan.FromMinutes(30);

        private static readonly Regex _loginConnexion = new(@"Connexion admin\s*:\s*(\S+)", RegexOptions.Compiled);
        private static readonly Regex _loginRefus = new(@"login\s*=\s*«\s*(.*?)\s*»", RegexOptions.Compiled);
        private static readonly Regex _parentheses = new(@"\s*\([^)]*\)", RegexOptions.Compiled);

        /// <param name="chronologique">Entrées du plus ancien au plus récent.</param>
        /// <returns>Blocs du plus récent au plus ancien (actions de chaque bloc dans l'ordre).</returns>
        public static List<BlocJournalAdmin> Regrouper(IEnumerable<EntreeJournalAdmin> chronologique)
        {
            var blocs = new List<BlocJournalAdmin>();
            var connexionParPoste = new Dictionary<string, BlocJournalAdmin>(StringComparer.OrdinalIgnoreCase);

            foreach (var e in chronologique)
            {
                string machine = e.Machine ?? string.Empty;

                if (string.Equals(e.Action, "ACCES_ADMIN_OK", StringComparison.OrdinalIgnoreCase))
                {
                    // Les anciennes lignes notaient l'opérateur : le vrai login admin est dans le détail.
                    var m = _loginConnexion.Match(e.Detail);
                    var bloc = new BlocJournalAdmin
                    {
                        Type = TypeBlocAdmin.Connexion,
                        Debut = e.Horodatage,
                        Utilisateur = m.Success ? m.Groups[1].Value : e.Utilisateur,
                        Machine = machine,
                    };
                    blocs.Add(bloc);
                    connexionParPoste[machine] = bloc;
                    continue;
                }

                if (string.Equals(e.Action, "ACCES_ADMIN_KO", StringComparison.OrdinalIgnoreCase))
                {
                    string login = _loginRefus.Match(e.Detail) is { Success: true } mr ? mr.Groups[1].Value : string.Empty;
                    var dernier = blocs.LastOrDefault(b => b.Type == TypeBlocAdmin.Refus
                        && string.Equals(b.Machine, machine, StringComparison.OrdinalIgnoreCase));
                    if (dernier != null && blocs[^1] == dernier && e.Horodatage - dernier.Debut <= FenetreRefus)
                    {
                        dernier.NbTentatives++;
                        if (login.Length > 0 && !dernier.LoginsTentes.Split(", ").Contains(login))
                            dernier.LoginsTentes = dernier.LoginsTentes.Length == 0 ? login : $"{dernier.LoginsTentes}, {login}";
                    }
                    else
                    {
                        blocs.Add(new BlocJournalAdmin
                        {
                            Type = TypeBlocAdmin.Refus, Debut = e.Horodatage, Machine = machine,
                            Utilisateur = e.Utilisateur, NbTentatives = 1, LoginsTentes = login,
                        });
                    }
                    continue;
                }

                BlocJournalAdmin cible;
                if (JournalAdminService.EstActionAutomatique(e.Action))
                {
                    // Les mises à jour automatiques d'une même journée tiennent dans un seul bloc.
                    var dernier = blocs.Count > 0 ? blocs[^1] : null;
                    if (dernier is { Type: TypeBlocAdmin.Automatique } && dernier.Debut.Date == e.Horodatage.Date)
                        cible = dernier;
                    else
                        blocs.Add(cible = new BlocJournalAdmin
                        {
                            Type = TypeBlocAdmin.Automatique, Debut = e.Horodatage, Machine = machine,
                        });
                }
                else if (connexionParPoste.TryGetValue(machine, out var connexion)
                         && e.Horodatage - connexion.Debut <= DureeMaxSession)
                {
                    cible = connexion;
                }
                else
                {
                    var dernier = blocs.Count > 0 ? blocs[^1] : null;
                    if (dernier is { Type: TypeBlocAdmin.SansConnexion }
                        && string.Equals(dernier.Machine, machine, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(dernier.Utilisateur, e.Utilisateur, StringComparison.OrdinalIgnoreCase)
                        && dernier.Debut.Date == e.Horodatage.Date)
                        cible = dernier;
                    else
                        blocs.Add(cible = new BlocJournalAdmin
                        {
                            Type = TypeBlocAdmin.SansConnexion, Debut = e.Horodatage,
                            Machine = machine, Utilisateur = e.Utilisateur,
                        });
                }

                AjouterAction(cible, e);
            }

            blocs.Reverse();
            return blocs;
        }

        /// <summary>Même action que la précédente du bloc (à un détail entre parenthèses près, comme
        /// un nombre de lignes) : on incrémente au lieu d'ajouter une ligne.</summary>
        private static void AjouterAction(BlocJournalAdmin bloc, EntreeJournalAdmin e)
        {
            var precedente = bloc.Actions.Count > 0 ? bloc.Actions[^1] : null;
            if (precedente != null
                && string.Equals(precedente.Code, e.Action, StringComparison.OrdinalIgnoreCase)
                && _parentheses.Replace(precedente.Detail, "") == _parentheses.Replace(e.Detail, ""))
            {
                precedente.Nombre++;
                precedente.Detail = e.Detail;
                return;
            }

            bloc.Actions.Add(new ActionAdmin
            {
                Heure = e.Horodatage,
                Code = e.Action,
                Libelle = e.ActionLisible,
                Detail = e.Detail,
            });
        }
    }
}
