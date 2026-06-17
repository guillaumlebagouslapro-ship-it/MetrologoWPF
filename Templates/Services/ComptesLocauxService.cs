using Metrologo.Models;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// CRUD des comptes utilisateurs (JSON via <see cref="Preferences"/>, pas de SQL).
    /// Identification à deux niveaux :
    ///  - Dropdown démarrage : déclaratif (pas de mdp), pour le journal et l'affichage.
    ///  - Modale Administration : login + mdp ; le rôle débloqué dépend du compte authentifié.
    /// Admin/SuperAdmin ont un mdp. Promotion → mdp généré ; rétrogradation → mdp effacé.
    /// </summary>
    public static class ComptesLocauxService
    {
        public static IReadOnlyList<Utilisateur> Lister()
        {
            AssurerCompteParDefaut();
            return Preferences.Utilisateurs
                .OrderBy(u => u.Nom)
                .ThenBy(u => u.Prenom)
                .ToList();
        }

        /// <summary>Crée un compte démo SuperAdmin (mdp « admin ») si la liste est vide.
        /// Supprimable/renommable depuis l'admin à tout moment.</summary>
        public static void AssurerCompteParDefaut()
        {
            if (Preferences.Utilisateurs.Count > 0)
            {
                AssurerAuMoinsUnSuperAdmin();
                AssurerPasswordsAdmins();
                return;
            }

            var defaut = new Utilisateur
            {
                Id = 1,
                Login = "utilisateur.demo",
                Nom = "Démo",
                Prenom = "Utilisateur",
                Role = RoleUtilisateur.SuperAdministrateur,
                Actif = true,
                DateCreation = DateTime.Now,
                PasswordHash = PasswordHasher.HashPassword(MotDePasseDefautDemo),
            };
            Preferences.SauvegarderUtilisateurs(new[] { defaut });

            JournalLog.Info(CategorieLog.Administration, "COMPTE_DEMO_CREE",
                "Compte de démonstration créé : utilisateur.demo (super-administrateur, mdp = « admin »).");
        }

        /// <summary>Mot de passe initial du compte Démo. À changer après le 1er login admin.</summary>
        public const string MotDePasseDefautDemo = "admin";

        // -------- Lecture / écriture compte --------

        public static Utilisateur Ajouter(string nom, string prenom)
        {
            if (string.IsNullOrWhiteSpace(nom)) throw new ArgumentException("Nom requis", nameof(nom));
            if (string.IsNullOrWhiteSpace(prenom)) throw new ArgumentException("Prénom requis", nameof(prenom));

            var liste = Preferences.Utilisateurs.ToList();
            string baseLogin = LoginGenerator.GenererBase(prenom, nom);
            string login = TrouverLoginDisponible(liste, baseLogin);

            int nextId = liste.Count == 0 ? 1 : liste.Max(u => u.Id) + 1;
            var nouveau = new Utilisateur
            {
                Id = nextId,
                Login = login,
                Nom = nom.Trim(),
                Prenom = prenom.Trim(),
                Role = RoleUtilisateur.Utilisateur,    // pas d'admin par défaut → pas de mdp
                Actif = true,
                DateCreation = DateTime.Now,
                PasswordHash = null,
            };

            liste.Add(nouveau);
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_AJOUTE",
                $"Compte ajouté : {login} ({prenom} {nom}).");
            return nouveau;
        }

        public static bool Supprimer(int id)
        {
            var liste = Preferences.Utilisateurs.ToList();
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null) return false;

            if (cible.Role == RoleUtilisateur.SuperAdministrateur && EstDernierSuperAdmin(id))
            {
                throw new InvalidOperationException(
                    "Impossible de supprimer le dernier super-administrateur. "
                  + "Promeus d'abord un autre compte en super-administrateur.");
            }

            liste.Remove(cible);
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_SUPPRIME",
                $"Compte supprimé : {cible.Login} (Id={id}).");
            return true;
        }

        public static bool Renommer(int id, string nom, string prenom)
        {
            if (string.IsNullOrWhiteSpace(nom)) throw new ArgumentException("Nom requis", nameof(nom));
            if (string.IsNullOrWhiteSpace(prenom)) throw new ArgumentException("Prénom requis", nameof(prenom));

            var liste = Preferences.Utilisateurs.ToList();
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null) return false;

            cible.Nom = nom.Trim();
            cible.Prenom = prenom.Trim();
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_RENOMME",
                $"Compte renommé : {cible.Login} → {prenom} {nom}.");
            return true;
        }

        // -------- Rôles + mots de passe --------

        /// <summary>Change le rôle. Retourne le mdp clair généré lors d'une promotion vers
        /// Admin/SuperAdmin, sinon null. Impossible de rétrograder le dernier SuperAdmin.</summary>
        public static string? ChangerRole(int id, RoleUtilisateur nouveauRole)
        {
            var liste = Preferences.Utilisateurs.ToList();
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null) return null;

            if (cible.Role == nouveauRole) return null;

            if (cible.Role == RoleUtilisateur.SuperAdministrateur
                && nouveauRole != RoleUtilisateur.SuperAdministrateur
                && EstDernierSuperAdmin(id))
            {
                throw new InvalidOperationException(
                    "Impossible de retirer le statut super-administrateur au dernier compte qui le possède. "
                  + "Promeus d'abord un autre compte en super-administrateur.");
            }

            string? mdpClair = null;
            bool deviendraAdmin = nouveauRole != RoleUtilisateur.Utilisateur;
            bool etaitUtilisateur = cible.Role == RoleUtilisateur.Utilisateur;

            if (nouveauRole == RoleUtilisateur.Utilisateur)
            {
                // Rétrogradation : mdp effacé.
                cible.PasswordHash = null;
            }
            else if (deviendraAdmin && (etaitUtilisateur || string.IsNullOrEmpty(cible.PasswordHash)))
            {
                // Promotion depuis Utilisateur OU compte admin sans mdp : on génère un mdp.
                mdpClair = GenererMotDePasse();
                cible.PasswordHash = PasswordHasher.HashPassword(mdpClair);
            }
            // Admin ↔ SuperAdmin : mdp existant conservé.

            cible.Role = nouveauRole;
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_ROLE",
                $"Rôle modifié : {cible.Login} → {nouveauRole}"
              + (mdpClair != null ? " (nouveau mdp généré)." : "."));
            return mdpClair;
        }

        /// <summary>Génère et enregistre un nouveau mdp pour un compte admin.
        /// Retourne le mdp clair (à afficher une seule fois). Lève si le compte est introuvable
        /// ou n'est pas admin.</summary>
        public static string ReinitialiserMotDePasse(int id)
        {
            var liste = Preferences.Utilisateurs.ToList();
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null)
                throw new InvalidOperationException($"Utilisateur Id={id} introuvable.");
            if (cible.Role == RoleUtilisateur.Utilisateur)
                throw new InvalidOperationException(
                    "Un utilisateur lambda n'a pas de mot de passe. Promeus-le d'abord en administrateur.");

            string mdpClair = GenererMotDePasse();
            cible.PasswordHash = PasswordHasher.HashPassword(mdpClair);
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Warn(CategorieLog.Administration, "MDP_REINITIALISE",
                $"Mot de passe réinitialisé pour {cible.Login}.");
            return mdpClair;
        }

        /// <summary>Définit un mdp choisi par l'admin (« Changer mon mdp »).
        /// L'ancien hash est remplacé immédiatement.</summary>
        public static bool DefinirMotDePasse(int id, string motDePasse)
        {
            if (string.IsNullOrWhiteSpace(motDePasse))
                throw new ArgumentException("Mot de passe vide", nameof(motDePasse));

            var liste = Preferences.Utilisateurs.ToList();
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null) return false;
            if (cible.Role == RoleUtilisateur.Utilisateur)
                throw new InvalidOperationException(
                    "Un utilisateur lambda n'a pas de mot de passe.");

            cible.PasswordHash = PasswordHasher.HashPassword(motDePasse);
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "MDP_MODIFIE",
                $"Mot de passe modifié pour {cible.Login}.");
            return true;
        }

        /// <summary>Authentifie un Admin/SuperAdmin (login + mdp). Retourne le compte si OK,
        /// null sinon. Ne distingue pas « login inconnu » de « mauvais mdp » (anti-énumération).</summary>
        public static Utilisateur? AuthentifierAdmin(string login, string motDePasse)
        {
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrEmpty(motDePasse)) return null;

            var cible = Preferences.Utilisateurs.FirstOrDefault(
                u => string.Equals(u.Login, login.Trim(), StringComparison.OrdinalIgnoreCase));
            if (cible == null) return null;
            if (cible.Role == RoleUtilisateur.Utilisateur) return null;
            if (!cible.Actif) return null;
            if (string.IsNullOrEmpty(cible.PasswordHash)) return null;
            if (!PasswordHasher.VerifyPassword(motDePasse, cible.PasswordHash)) return null;

            cible.DernierLogin = DateTime.Now;
            Preferences.SauvegarderUtilisateurs(Preferences.Utilisateurs);
            return cible;
        }

        // -------- Helpers --------

        /// <summary>Renvoie vrai si <paramref name="id"/> est le seul SuperAdmin du catalogue.</summary>
        public static bool EstDernierSuperAdmin(int id)
        {
            var liste = Preferences.Utilisateurs;
            var cible = liste.FirstOrDefault(u => u.Id == id);
            if (cible == null || cible.Role != RoleUtilisateur.SuperAdministrateur) return false;
            return liste.Count(u => u.Role == RoleUtilisateur.SuperAdministrateur) <= 1;
        }

        private static void AssurerAuMoinsUnSuperAdmin()
        {
            var liste = Preferences.Utilisateurs.ToList();
            if (liste.Any(u => u.Role == RoleUtilisateur.SuperAdministrateur)) return;
            if (liste.Count == 0) return;

            var premier = liste[0];
            premier.Role = RoleUtilisateur.SuperAdministrateur;
            Preferences.SauvegarderUtilisateurs(liste);

            JournalLog.Info(CategorieLog.Administration, "SUPER_ADMIN_AUTO",
                $"Aucun super-administrateur trouvé : promotion automatique de {premier.Login}.");
        }

        /// <summary>Migration : attribue le mdp par défaut « admin » aux comptes admin
        /// sans hash. Récupère les bases créées avant l'introduction des passwords.</summary>
        private static void AssurerPasswordsAdmins()
        {
            var liste = Preferences.Utilisateurs.ToList();
            bool modifie = false;
            foreach (var u in liste)
            {
                if (u.Role != RoleUtilisateur.Utilisateur && string.IsNullOrEmpty(u.PasswordHash))
                {
                    u.PasswordHash = PasswordHasher.HashPassword(MotDePasseDefautDemo);
                    modifie = true;
                    JournalLog.Warn(CategorieLog.Administration, "MDP_DEFAUT_FORCE",
                        $"Mot de passe par défaut « admin » attribué à {u.Login} (migration).");
                }
            }
            if (modifie) Preferences.SauvegarderUtilisateurs(liste);
        }

        private static string TrouverLoginDisponible(IList<Utilisateur> liste, string baseLogin)
        {
            string candidat = baseLogin;
            for (int suffixe = 2; suffixe < 1000; suffixe++)
            {
                if (!liste.Any(u => string.Equals(u.Login, candidat, StringComparison.OrdinalIgnoreCase)))
                    return candidat;
                candidat = $"{baseLogin}{suffixe}";
            }
            return $"{baseLogin}{Guid.NewGuid().ToString("N")[..6]}";
        }

        /// <summary>Génère un mdp aléatoire lisible (8 car. par défaut, lettres + chiffres,
        /// pas de symboles). Entropie via <see cref="RandomNumberGenerator"/>.</summary>
        public static string GenererMotDePasse(int longueur = 8)
        {
            const string alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
            // Ambigus retirés : O/0/o, I/1/l/L
            var buf = new char[longueur];
            byte[] rand = RandomNumberGenerator.GetBytes(longueur);
            for (int i = 0; i < longueur; i++)
                buf[i] = alphabet[rand[i] % alphabet.Length];
            return new string(buf);
        }
    }
}
