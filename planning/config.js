// Réglages de SIRØB Planning.
// - clientId : l'app Azure « SIRØB Project Browser » (la même que SIRØB Mobile et le Mac).
// - googleClientId : l'identifiant OAuth Google créé pour lire ton agenda (voir README, étape 3).
//   Tant qu'il est vide, l'app marche sans l'agenda.
// Aucun de ces identifiants n'est un secret.
window.SIROB_CONFIG = {
  clientId: "3a35d60e-a1e1-4b40-8f59-ac00fb3077de",
  catalogue: "SIROB Browser/catalogue.json",
  planning: "SIROB Browser/planning.json",
  googleClientId: "1045337931059-6ev1vuhjmn3u536967i0rt163j22lft9.apps.googleusercontent.com",
};
