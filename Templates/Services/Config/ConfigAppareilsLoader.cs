using System;
using System.Collections.Generic;
using System.IO;
using Metrologo.Models;

namespace Metrologo.Services.Config
{
    /// <summary>
    /// Charge <see cref="ConfigAppareils"/> depuis Metrologo.ini.
    /// Portage direct de la logique Delphi F_Main.ChargementParametres (F_Main.pas:540).
    /// </summary>
    public static class ConfigAppareilsLoader
    {
        public const string SectionStanford = "Stanford SR620";
        public const string SectionRacal    = "Racal-Dana 1996";
        public const string SectionEip      = "EIP 545";
        public const string SectionMux      = "HP59307A";

        // Correspondance Gate{lib} → index utilisé dans EnTetesMesureHelper (0..12).
        // Aligné sur le tableau _libellesGate de EnTetesMesure.cs.
        private static readonly (string Cle, int Index, double Secondes)[] _gates =
        {
            ("Gate10ms",  0, 0.010),
            ("Gate20ms",  1, 0.020),
            ("Gate50ms",  2, 0.050),
            ("Gate100ms", 3, 0.100),
            ("Gate200ms", 4, 0.200),
            ("Gate500ms", 5, 0.500),
            ("Gate1s",    6, 1.0),
            ("Gate2s",    7, 2.0),
            ("Gate5s",    8, 5.0),
            ("Gate10s",   9, 10.0),
            ("Gate20s",  10, 20.0),
            ("Gate50s",  11, 50.0),
            ("Gate100s", 12, 100.0)
        };

        public static ConfigAppareils Charger(string cheminFichier)
        {
            var ini = IniFile.Charger(cheminFichier);
            var avertissements = new List<string>();

            var stanford = LireAppareil(ini, SectionStanford, avecSrqObligatoire: true);
            var racal    = LireAppareil(ini, SectionRacal,    avecSrqObligatoire: true);
            var eip      = LireAppareil(ini, SectionEip,      avecSrqObligatoire: true);

            AppareilIEEE? mux = null;
            if (ini.ContientSection(SectionMux))
            {
                try { mux = LireMux(ini, SectionMux); }
                catch (Exception ex)
                {
                    avertissements.Add(
                        $"Multiplexeur « {SectionMux} » : erreur de chargement ({ex.Message}). "
                        + "Démarrage sans multiplexeur.");
                }
            }
            else
            {
                avertissements.Add(
                    $"Multiplexeur « {SectionMux} » absent de la configuration. "
                    + "Démarrage sans multiplexeur (fonctionnalité non utilisée aujourd'hui).");
            }

            // Détection du conflit d'adresses GPIB historique (Stanford vs EIP à l'adresse 16).
            DetecterConflitsAdresses(avertissements, stanford, racal, eip, mux);

            return new ConfigAppareils
            {
                Stanford = stanford,
                Racal = racal,
                Eip = eip,
                Mux = mux,
                Avertissements = avertissements
            };
        }

        private static AppareilIEEE LireAppareil(IniFile ini, string section, bool avecSrqObligatoire)
        {
            if (!ini.ContientSection(section))
                throw new InvalidDataException(
                    $"Section [{section}] introuvable dans Metrologo.ini.");

            var app = new AppareilIEEE
            {
                Nom = section,
                Adresse = ini.EntierObligatoire(section, "Adresse"),
                WriteTerm = ini.EntierObligatoire(section, "WriteTerm"),
                ReadTerm = ini.EntierObligatoire(section, "ReadTerm"),
                TailleHeaderReponse = ini.EntierObligatoire(section, "TailleHeaderReponse"),
                ChaineInit = ini.Valeur(section, "ChaineInit") ?? string.Empty,
                ConfEntree = ini.Valeur(section, "ConfEntreeDef") ?? string.Empty,
                ExeMesure = ini.Valeur(section, "ExeMesure") ?? string.Empty,
                Monocoup = ini.Valeur(section, "Monocoup") ?? string.Empty,
                GereSRQ = avecSrqObligatoire && ini.BooleenObligatoire(section, "GestSRQ"),
                SRQOn = ini.Valeur(section, "SRQOn") ?? string.Empty,
                SRQOff = ini.Valeur(section, "SRQOff") ?? string.Empty
            };

            foreach (var (cle, index, secondes) in _gates)
            {
                var cmd = ini.Valeur(section, cle);
                if (string.IsNullOrWhiteSpace(cmd)) continue;  // Gates vides ignorées (cf. AffectationValeur Delphi)
                app.Gates[index] = new GateConfig
                {
                    Libelle = LibelleGate(index),
                    Commande = cmd,
                    ValeurSecondes = secondes
                };
            }

            return app;
        }

        private static AppareilIEEE LireMux(IniFile ini, string section) => new()
        {
            Nom = section,
            Adresse = ini.EntierObligatoire(section, "Adresse"),
            WriteTerm = ini.EntierObligatoire(section, "WriteTerm"),
            ReadTerm = ini.EntierObligatoire(section, "ReadTerm")
        };

        private static void DetecterConflitsAdresses(
            List<string> avertissements, params AppareilIEEE?[] appareils)
        {
            var parAdresse = new Dictionary<int, List<string>>();
            foreach (var a in appareils)
            {
                if (a == null) continue;
                if (!parAdresse.TryGetValue(a.Adresse, out var liste))
                {
                    liste = new List<string>();
                    parAdresse[a.Adresse] = liste;
                }
                liste.Add(a.Nom);
            }

            foreach (var (adresse, noms) in parAdresse)
            {
                if (noms.Count > 1)
                    avertissements.Add(
                        $"Adresse GPIB {adresse} partagée par {noms.Count} appareils "
                        + $"({string.Join(", ", noms)}) : un seul peut être branché à la fois.");
            }
        }

        private static string LibelleGate(int index) => index switch
        {
            0 => "10 ms", 1 => "20 ms", 2 => "50 ms",
            3 => "100 ms", 4 => "200 ms", 5 => "500 ms",
            6 => "1 s", 7 => "2 s", 8 => "5 s",
            9 => "10 s", 10 => "20 s", 11 => "50 s", 12 => "100 s",
            _ => ""
        };
    }
}
