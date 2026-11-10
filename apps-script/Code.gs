/**
 * Prime League FC 27 : formulaires de déclaration et de validation des matchs.
 *
 * Installation (une seule fois) :
 *  1. Dans le Google Sheet : Extensions > Apps Script.
 *  2. Remplacer le contenu de Code.gs par ce fichier, puis enregistrer.
 *  3. Recharger l'onglet du Google Sheet : un menu "Prime League" apparaît.
 *  4. Cliquer sur Prime League > Installer les formulaires et accepter les
 *     autorisations demandées par Google (relancer le menu si besoin).
 *  5. Les liens des formulaires s'affichent.
 *
 * Fonctionnement :
 *  - Un joueur déclare un match avec le formulaire "Déclarer un match".
 *  - Le match est ajouté dans l'onglet Matchs avec le statut "En attente".
 *  - L'adversaire reçoit un email avec un lien vers le formulaire
 *    "Valider un match", déjà rempli avec le code du match.
 *  - S'il confirme, le statut passe à "Validé". S'il conteste, il passe à
 *    "Contesté" et l'organisateur ainsi que le déclarant sont prévenus.
 *  - L'organisateur peut toujours corriger un statut à la main dans l'onglet Matchs.
 *  - La colonne M "Email de validation" indique si l'email est parti ou l'erreur
 *    rencontrée. Le menu Prime League permet de renvoyer les emails des matchs
 *    en attente et de lancer un diagnostic (déclencheurs, quota, email de test).
 */

var ONGLET_JOUEURS = 'Joueurs';
var ONGLET_MATCHS = 'Matchs';
var ONGLET_PARAMETRES = 'Paramètres';
var ONGLET_REP_DECLARATIONS = 'Réponses déclarations';
var ONGLET_REP_VALIDATIONS = 'Réponses validations';

var Q_JOUEUR = 'Ton nom';
var Q_ADVERSAIRE = 'Ton adversaire';
var Q_ETOILES_JOUEUR = 'Étoiles de ton équipe';
var Q_ETOILES_ADVERSAIRE = "Étoiles de l'équipe adverse";
var Q_BUTS_JOUEUR = 'Tes buts';
var Q_BUTS_ADVERSAIRE = "Buts de l'adversaire";

var Q_CODE = 'Code du match';
var Q_DECISION = 'Ta décision';
var Q_COMMENTAIRE = 'Commentaire (facultatif)';
var CHOIX_CONFIRME = 'Je confirme le résultat';
var CHOIX_CONTESTE = 'Je conteste le résultat';

var NOM_EXPEDITEUR = 'PrimeLeague Validator';

var STATUT_ATTENTE = 'En attente';
var STATUT_VALIDE = 'Validé';
var STATUT_CONTESTE = 'Contesté';
var STATUT_REJETE = 'Rejeté';

var CHOIX_ETOILES = ['0,5', '1', '1,5', '2', '2,5', '3', '3,5', '4', '4,5', '5'];

// Colonnes de l'onglet Matchs (1 = A)
var COL = {
  ID: 1, DECLARE_LE: 2, J1: 3, ETOILES_J1: 4, BUTS_J1: 5,
  J2: 6, ETOILES_J2: 7, BUTS_J2: 8, STATUT: 9, MAJ_LE: 10,
  COMMENTAIRE: 11, CODE: 12, EMAIL: 13
};
var NB_COLONNES_MATCHS = 12;
var TITRE_COL_EMAIL = 'Email de validation';

/* ------------------------------------------------------------------ */
/* Menu                                                                */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Prime League')
    .addItem('Installer les formulaires', 'installer')
    .addItem('Mettre à jour la liste des joueurs', 'majListeJoueurs')
    .addItem('Afficher les liens des formulaires', 'afficherLiens')
    .addItem('Réinstaller le déclencheur', 'envoyerDepuisMonCompte')
    .addItem('Rattraper les déclarations manquantes', 'rattraperDeclarations')
    // Emails de validation en pause : menus à réactiver avec l'envoi.
    // .addSeparator()
    // .addItem('Renvoyer les emails des matchs en attente', 'renvoyerEmailsEnAttente')
    // .addItem('Diagnostic des emails', 'diagnosticEmails')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* Installation                                                        */
/* ------------------------------------------------------------------ */

function installer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();

  if (props.getProperty('FORM_DECLARATION_ID')) {
    SpreadsheetApp.getUi().alert(
      'Les formulaires sont déjà installés. Utilise le menu Prime League > Afficher les liens des formulaires.');
    return;
  }

  // Formulaire de déclaration
  var fDecl = FormApp.create('Prime League : déclarer un match');
  fDecl.setDescription(
    "Déclare le résultat d'un match 1v1. Ton adversaire recevra un email pour le confirmer. " +
    "Le match ne compte au classement qu'une fois confirmé.");
  fDecl.setConfirmationMessage("Merci ! Ton adversaire va recevoir un email pour confirmer le résultat.");
  fDecl.addListItem().setTitle(Q_JOUEUR).setRequired(true);
  fDecl.addListItem().setTitle(Q_ADVERSAIRE).setRequired(true);
  fDecl.addListItem().setTitle(Q_ETOILES_JOUEUR)
    .setHelpText("Le nombre d'étoiles affiché par FC 27 pour le club que tu as joué.")
    .setChoiceValues(CHOIX_ETOILES).setRequired(true);
  fDecl.addListItem().setTitle(Q_ETOILES_ADVERSAIRE)
    .setChoiceValues(CHOIX_ETOILES).setRequired(true);
  var butsValidation = FormApp.createTextValidation()
    .setHelpText('Un nombre entier, 0 ou plus.')
    .requireWholeNumber()
    .build();
  fDecl.addTextItem().setTitle(Q_BUTS_JOUEUR).setValidation(butsValidation).setRequired(true);
  fDecl.addTextItem().setTitle(Q_BUTS_ADVERSAIRE).setValidation(butsValidation).setRequired(true);
  fDecl.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Formulaire de validation
  var fVal = FormApp.create('Prime League : valider un match');
  fVal.setDescription("Confirme ou conteste le résultat déclaré par ton adversaire.");
  fVal.setConfirmationMessage('Merci, ta réponse a été enregistrée.');
  fVal.addTextItem().setTitle(Q_CODE)
    .setHelpText("Rempli automatiquement depuis le lien reçu par email. Ne le modifie pas.")
    .setRequired(true);
  fVal.addMultipleChoiceItem().setTitle(Q_DECISION)
    .setChoiceValues([CHOIX_CONFIRME, CHOIX_CONTESTE]).setRequired(true);
  fVal.addParagraphTextItem().setTitle(Q_COMMENTAIRE)
    .setHelpText('Si tu contestes, indique le bon score ou ce qui ne va pas.');
  fVal.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Renommer les onglets de réponses créés par Google Forms
  var ongletDecl = attendreOngletReponses_(ss, fDecl.getId());
  var ongletVal = attendreOngletReponses_(ss, fVal.getId());
  ongletDecl.setName(ONGLET_REP_DECLARATIONS);
  ongletVal.setName(ONGLET_REP_VALIDATIONS);
  ss.setActiveSheet(ongletDecl); ss.moveActiveSheet(ss.getNumSheets());
  ss.setActiveSheet(ongletVal); ss.moveActiveSheet(ss.getNumSheets());
  ss.setActiveSheet(ss.getSheetByName(ONGLET_MATCHS));

  props.setProperties({
    FORM_DECLARATION_ID: fDecl.getId(),
    FORM_VALIDATION_ID: fVal.getId(),
    ONGLET_DECL_ID: String(ongletDecl.getSheetId()),
    ONGLET_VAL_ID: String(ongletVal.getSheetId())
  });

  creerDeclencheurs_(ss);

  majListeJoueurs();
  afficherLiens();
}

// Déclencheurs : réponse à un formulaire, modification de la liste des joueurs.
// Un déclencheur s'exécute avec le compte Google qui l'a créé : c'est ce compte
// qui envoie les emails. getProjectTriggers ne renvoie que les déclencheurs du
// compte connecté.
function creerDeclencheurs_(ss) {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('surEnvoiFormulaire').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('surModification').forSpreadsheet(ss).onEdit().create();
  // Filet de sécurité : toutes les heures, reprend les déclarations oubliées.
  ScriptApp.newTrigger('rattraperDeclarationsAuto').timeBased().everyHours(1).create();
}

/**
 * Fait envoyer les emails par le compte connecté (par exemple l'adresse
 * professionnelle), sans recréer les formulaires. Il faut ensuite supprimer
 * les déclencheurs de l'ancien compte, sinon les emails partent en double.
 */
function envoyerDepuisMonCompte() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var moi = Session.getEffectiveUser().getEmail();
  try {
    FormApp.openById(props.getProperty('FORM_DECLARATION_ID'));
    FormApp.openById(props.getProperty('FORM_VALIDATION_ID'));
  } catch (err) {
    ui.alert("Le compte " + moi + " n'a pas accès aux formulaires. Le propriétaire doit d'abord " +
      "l'ajouter comme collaborateur sur les deux formulaires (Prime League > Afficher les liens des formulaires).\n\n" +
      'Détail : ' + err.message);
    return;
  }
  var ok = ui.alert('Prime League',
    'Le déclencheur va être installé avec le compte ' + moi + '.\n\n' +
    "Un seul compte doit avoir les déclencheurs (Apps Script > Déclencheurs), " +
    'sinon chaque match sera ajouté deux fois. Continuer ?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  creerDeclencheurs_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert('C\'est fait : le déclencheur tourne maintenant avec ' + moi + '.');
}

function attendreOngletReponses_(ss, formId) {
  for (var essai = 0; essai < 20; essai++) {
    SpreadsheetApp.flush();
    var onglets = ss.getSheets();
    for (var i = 0; i < onglets.length; i++) {
      var url = onglets[i].getFormUrl();
      if (url && FormApp.openByUrl(url).getId() === formId) return onglets[i];
    }
    Utilities.sleep(1000);
  }
  throw new Error("L'onglet de réponses du formulaire n'a pas été trouvé. Relance l'installation.");
}

function afficherLiens() {
  var props = PropertiesService.getDocumentProperties();
  var idDecl = props.getProperty('FORM_DECLARATION_ID');
  if (!idDecl) {
    SpreadsheetApp.getUi().alert("Les formulaires ne sont pas encore installés. Lance d'abord Prime League > Installer les formulaires.");
    return;
  }
  var fDecl = FormApp.openById(idDecl);
  var fVal = FormApp.openById(props.getProperty('FORM_VALIDATION_ID'));
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px;line-height:1.5">' +
    '<p><b>Lien à partager aux joueurs (déclarer un match) :</b><br>' +
    '<a href="' + fDecl.getPublishedUrl() + '" target="_blank">' + fDecl.getPublishedUrl() + '</a></p>' +
    '<p><b>Modifier le formulaire de déclaration :</b><br>' +
    '<a href="' + fDecl.getEditUrl() + '" target="_blank">ouvrir</a></p>' +
    '<p><b>Modifier le formulaire de validation :</b><br>' +
    '<a href="' + fVal.getEditUrl() + '" target="_blank">ouvrir</a><br>' +
    "<i>Pas besoin de partager celui-ci : chaque adversaire reçoit son lien par email.</i></p>" +
    '</div>').setWidth(520).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Prime League : formulaires');
}

/* ------------------------------------------------------------------ */
/* Liste des joueurs                                                   */
/* ------------------------------------------------------------------ */

function lireJoueurs_() {
  var onglet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_JOUEURS);
  var der = onglet.getLastRow();
  if (der < 2) return [];
  return onglet.getRange(2, 1, der - 1, 2).getValues()
    .map(function (l) { return { nom: String(l[0]).trim(), email: String(l[1]).trim() }; })
    .filter(function (j) { return j.nom !== ''; });
}

function majListeJoueurs() {
  var idDecl = PropertiesService.getDocumentProperties().getProperty('FORM_DECLARATION_ID');
  if (!idDecl) return;
  var noms = lireJoueurs_().map(function (j) { return j.nom; })
    .filter(function (n, i, t) { return t.indexOf(n) === i; })
    .sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  if (noms.length === 0) noms = ['(aucun joueur dans l\'onglet Joueurs)'];
  var form = FormApp.openById(idDecl);
  form.getItems(FormApp.ItemType.LIST).forEach(function (item) {
    var titre = item.getTitle();
    if (titre === Q_JOUEUR || titre === Q_ADVERSAIRE) {
      item.asListItem().setChoiceValues(noms);
    }
  });
}

function surModification(e) {
  if (e && e.range && e.range.getSheet().getName() === ONGLET_JOUEURS) {
    majListeJoueurs();
  }
}

/* ------------------------------------------------------------------ */
/* Réception des formulaires                                           */
/* ------------------------------------------------------------------ */

function surEnvoiFormulaire(e) {
  var props = PropertiesService.getDocumentProperties();
  var idOnglet = String(e.range.getSheet().getSheetId());
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (idOnglet === props.getProperty('ONGLET_DECL_ID')) {
      traiterDeclaration_(e.namedValues);
    } else if (idOnglet === props.getProperty('ONGLET_VAL_ID')) {
      traiterValidation_(e.namedValues);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Rattrapage des déclarations absentes de l'onglet Matchs             */
/* ------------------------------------------------------------------ */

function rattraperDeclarations() {
  var n = rattraperDeclarations_();
  SpreadsheetApp.getUi().alert(n ? n + ' déclaration(s) ajoutée(s) dans Matchs.'
    : 'Rien à rattraper : toutes les déclarations sont dans Matchs.');
}

function rattraperDeclarationsAuto() {
  rattraperDeclarations_();
}

function dateDe_(v) {
  if (v instanceof Date) return v;
  var m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0)) : null;
}

/**
 * Ajoute dans Matchs chaque réponse du formulaire de déclaration qui n'y est
 * pas encore (même déclarant, adversaire, score, à 15 minutes près).
 * Renvoie le nombre de matchs ajoutés.
 */
function rattraperDeclarations_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  var idDecl = props.getProperty('ONGLET_DECL_ID');
  var rep = ss.getSheets().filter(function (o) { return String(o.getSheetId()) === idDecl; })[0] ||
    ss.getSheetByName(ONGLET_REP_DECLARATIONS);
  var matchs = ss.getSheetByName(ONGLET_MATCHS);
  if (!rep || rep.getLastRow() < 2) return 0;

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var donnees = rep.getRange(1, 1, rep.getLastRow(), rep.getLastColumn()).getDisplayValues();
    var dates = rep.getRange(1, 1, rep.getLastRow(), 1).getValues();
    var entetes = donnees[0];
    var existants = matchs.getLastRow() < 2 ? [] :
      matchs.getRange(2, 1, matchs.getLastRow() - 1, NB_COLONNES_MATCHS).getValues();
    var ajoutes = 0;
    for (var i = 1; i < donnees.length; i++) {
      var ligne = donnees[i];
      if (!ligne[0]) continue;
      var nv = {};
      entetes.forEach(function (t, k) { nv[t] = [ligne[k]]; });
      var quand = dateDe_(dates[i][0]) || dateDe_(ligne[0]);
      var j1 = valeur_(nv, Q_JOUEUR), j2 = valeur_(nv, Q_ADVERSAIRE);
      var b1 = Number(valeur_(nv, Q_BUTS_JOUEUR)), b2 = Number(valeur_(nv, Q_BUTS_ADVERSAIRE));
      var dejaLa = existants.some(function (l) {
        var d = dateDe_(l[COL.DECLARE_LE - 1]);
        return String(l[COL.J1 - 1]).trim() === j1 && String(l[COL.J2 - 1]).trim() === j2 &&
          Number(l[COL.BUTS_J1 - 1]) === b1 && Number(l[COL.BUTS_J2 - 1]) === b2 &&
          (!d || !quand || Math.abs(d - quand) < 15 * 60 * 1000);
      });
      // Une réponse de moins de 2 minutes est encore en cours de traitement
      // par le déclencheur du formulaire : on la laisse pour éviter un doublon.
      if (dejaLa || (quand && new Date() - quand < 2 * 60 * 1000)) continue;
      traiterDeclaration_(nv, quand);
      existants.push([null, quand, j1, null, b1, j2, null, b2]);
      ajoutes++;
    }
    return ajoutes;
  } finally {
    lock.releaseLock();
  }
}

function valeur_(namedValues, question) {
  var v = namedValues[question];
  return v && v.length ? String(v[0]).trim() : '';
}

function etoiles_(texte) {
  return Number(String(texte).replace(',', '.'));
}

function traiterDeclaration_(nv, dateDeclaration) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var onglet = ss.getSheetByName(ONGLET_MATCHS);
  var joueurs = lireJoueurs_();
  var j1 = valeur_(nv, Q_JOUEUR);
  var j2 = valeur_(nv, Q_ADVERSAIRE);
  var e1 = etoiles_(valeur_(nv, Q_ETOILES_JOUEUR));
  var e2 = etoiles_(valeur_(nv, Q_ETOILES_ADVERSAIRE));
  var b1 = Number(valeur_(nv, Q_BUTS_JOUEUR));
  var b2 = Number(valeur_(nv, Q_BUTS_ADVERSAIRE));

  // Numéro suivant le plus grand ID existant (robuste aux lignes ajoutées à la main)
  var numero = 1;
  if (onglet.getLastRow() > 1) {
    onglet.getRange(2, COL.ID, onglet.getLastRow() - 1, 1).getValues().forEach(function (l) {
      var n = Number(String(l[0]).replace(/^M/, ''));
      if (n >= numero) numero = n + 1;
    });
  }
  var id = 'M' + ('000' + numero).slice(-4);
  var code = id + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
  var maintenant = new Date();
  var declareLe = dateDeclaration || maintenant;

  var adversaire = joueurs.filter(function (j) { return j.nom === j2; })[0];
  var statut = STATUT_ATTENTE;
  var commentaire = '';
  if (j1 === j2) {
    statut = STATUT_REJETE;
    commentaire = 'Rejeté automatiquement : le joueur et son adversaire sont identiques.';
  } else if (!adversaire || !adversaire.email) {
    commentaire = "Email de l'adversaire manquant dans l'onglet Joueurs : validation à faire par l'organisateur.";
  }

  onglet.appendRow([id, declareLe, j1, e1, b1, j2, e2, b2, statut, maintenant, commentaire, code]);
  var ligne = onglet.getLastRow();

  if (statut !== STATUT_ATTENTE) return;

  // EMAIL DE VALIDATION EN PAUSE (sécurité Apps Script côté Primelis) :
  // l'organisateur passe le statut à "Validé" à la main dans l'onglet Matchs.
  // Pour réactiver l'envoi, supprimer les "//" des 4 lignes ci-dessous et
  // supprimer la ligne noterEnvoi_ qui suit.
  // var resultat = envoyerEmailValidation_({
  //   id: id, j1: j1, j2: j2, e1: e1, e2: e2, b1: b1, b2: b2, code: code
  // }, adversaire);
  // noterEnvoi_(onglet, ligne, resultat);
  noterEnvoi_(onglet, ligne, 'Emails désactivés : validation manuelle');
}

/**
 * Envoie l'email de validation à l'adversaire (ou à l'organisateur si
 * l'adversaire n'a pas d'email). Ne lève jamais d'erreur : renvoie un texte
 * qui décrit le résultat, écrit ensuite dans la colonne "Email de validation".
 */
function envoyerEmailValidation_(m, adversaire) {
  var resume = m.j1 + ' ' + m.b1 + ' - ' + m.b2 + ' ' + m.j2 +
    ' (étoiles : ' + m.j1 + ' ' + String(m.e1).replace('.', ',') +
    ', ' + m.j2 + ' ' + String(m.e2).replace('.', ',') + ')';
  var horodatage = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  try {
    if (adversaire && adversaire.email) {
      var lien = lienValidation_(m.code);
      MailApp.sendEmail({
        name: NOM_EXPEDITEUR,
        to: adversaire.email,
        replyTo: emailOrganisateur_(),
        subject: 'Prime League : confirme ton match contre ' + m.j1,
        body:
          'Salut ' + m.j2 + ',\n\n' +
          m.j1 + ' a déclaré le résultat suivant :\n' + resume + '\n\n' +
          'Pour confirmer ou contester ce résultat, ouvre ce lien :\n' + lien + '\n\n' +
          "Le match comptera au classement dès que tu l'auras confirmé.",
        htmlBody:
          '<p>Salut ' + echapper_(m.j2) + ',</p>' +
          '<p>' + echapper_(m.j1) + ' a déclaré le résultat suivant :</p>' +
          '<p style="font-size:16px"><b>' + echapper_(resume) + '</b></p>' +
          '<p><a href="' + lien + '">Confirmer ou contester ce résultat</a></p>' +
          '<p>Le match comptera au classement dès que tu l\'auras confirmé.</p>'
      });
      return 'Envoyé à ' + adversaire.email + ' le ' + horodatage;
    }
    MailApp.sendEmail({
      name: NOM_EXPEDITEUR,
      to: emailOrganisateur_(),
      subject: 'Prime League : match ' + m.id + ' sans email adversaire',
      htmlBody: '<p>Le match ' + m.id + ' (' + echapper_(resume) + ") n'a pas pu être envoyé en validation, " +
        "car l'adversaire n'a pas d'email dans l'onglet Joueurs. Tu peux le valider à la main dans l'onglet Matchs.</p>"
    });
    return "Pas d'email adversaire : organisateur prévenu le " + horodatage;
  } catch (err) {
    console.error('Envoi impossible pour ' + m.id + ' : ' + err);
    return 'ÉCHEC le ' + horodatage + ' : ' + (err && err.message ? err.message : err);
  }
}

function noterEnvoi_(onglet, ligne, texte) {
  var entete = onglet.getRange(1, COL.EMAIL);
  if (!entete.getValue()) entete.setValue(TITRE_COL_EMAIL);
  onglet.getRange(ligne, COL.EMAIL).setValue(texte);
}

function lienValidation_(code) {
  var form = FormApp.openById(PropertiesService.getDocumentProperties().getProperty('FORM_VALIDATION_ID'));
  var itemCode = form.getItems(FormApp.ItemType.TEXT).filter(function (i) {
    return i.getTitle() === Q_CODE;
  })[0];
  return form.createResponse()
    .withItemResponse(itemCode.asTextItem().createResponse(code))
    .toPrefilledUrl();
}

function traiterValidation_(nv) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var onglet = ss.getSheetByName(ONGLET_MATCHS);
  var code = valeur_(nv, Q_CODE).toUpperCase();
  var decision = valeur_(nv, Q_DECISION);
  var commentaire = valeur_(nv, Q_COMMENTAIRE);

  var der = onglet.getLastRow();
  if (der < 2) return;
  var lignes = onglet.getRange(2, 1, der - 1, NB_COLONNES_MATCHS).getValues();
  for (var i = 0; i < lignes.length; i++) {
    if (String(lignes[i][COL.CODE - 1]).toUpperCase() !== code) continue;
    // On ne traite qu'une seule réponse par match
    if (lignes[i][COL.STATUT - 1] !== STATUT_ATTENTE) return;

    var ligne = i + 2;
    var statut = decision === CHOIX_CONTESTE ? STATUT_CONTESTE : STATUT_VALIDE;
    onglet.getRange(ligne, COL.STATUT, 1, 3).setValues([[statut, new Date(), commentaire]]);

    if (statut === STATUT_CONTESTE) {
      var l = lignes[i];
      var resume = l[COL.J1 - 1] + ' ' + l[COL.BUTS_J1 - 1] + ' - ' + l[COL.BUTS_J2 - 1] + ' ' + l[COL.J2 - 1];
      var declarant = lireJoueurs_().filter(function (j) { return j.nom === l[COL.J1 - 1]; })[0];
      var destinataires = [emailOrganisateur_()];
      if (declarant && declarant.email) destinataires.push(declarant.email);
      MailApp.sendEmail({
        name: NOM_EXPEDITEUR,
        to: destinataires.join(','),
        subject: 'Prime League : le match ' + l[COL.ID - 1] + ' est contesté',
        htmlBody: '<p>' + echapper_(l[COL.J2 - 1]) + ' conteste le résultat <b>' + echapper_(resume) + '</b>.</p>' +
          '<p>Commentaire : ' + (commentaire ? echapper_(commentaire) : '(aucun)') + '</p>' +
          "<p>L'organisateur peut corriger le score et passer le statut à Validé dans l'onglet Matchs, " +
          'ou le déclarant peut déclarer à nouveau le match avec le bon score.</p>'
      });
    }
    return;
  }
}

/* ------------------------------------------------------------------ */
/* Renvoi et diagnostic des emails                                     */
/* ------------------------------------------------------------------ */

function renvoyerEmailsEnAttente() {
  var ui = SpreadsheetApp.getUi();
  var onglet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_MATCHS);
  var der = onglet.getLastRow();
  if (der < 2) { ui.alert('Aucun match dans l\'onglet Matchs.'); return; }
  var lignes = onglet.getRange(2, 1, der - 1, NB_COLONNES_MATCHS).getValues();
  var joueurs = lireJoueurs_();
  var bilan = [];
  for (var i = 0; i < lignes.length; i++) {
    var l = lignes[i];
    if (l[COL.STATUT - 1] !== STATUT_ATTENTE || !l[COL.CODE - 1]) continue;
    var j2 = String(l[COL.J2 - 1]).trim();
    var adversaire = joueurs.filter(function (j) { return j.nom === j2; })[0];
    var resultat = envoyerEmailValidation_({
      id: l[COL.ID - 1], j1: l[COL.J1 - 1], j2: j2,
      e1: l[COL.ETOILES_J1 - 1], e2: l[COL.ETOILES_J2 - 1],
      b1: l[COL.BUTS_J1 - 1], b2: l[COL.BUTS_J2 - 1], code: l[COL.CODE - 1]
    }, adversaire);
    noterEnvoi_(onglet, i + 2, resultat);
    bilan.push(l[COL.ID - 1] + ' : ' + resultat);
  }
  ui.alert(bilan.length ? 'Emails renvoyés :\n\n' + bilan.join('\n') : 'Aucun match en attente.');
}

function diagnosticEmails() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lignes = [];

  lignes.push('Compte qui envoie les emails : ' + Session.getEffectiveUser().getEmail());
  lignes.push('Emails encore autorisés aujourd\'hui : ' + MailApp.getRemainingDailyQuota());

  var declencheurs = ScriptApp.getProjectTriggers().map(function (t) {
    return t.getHandlerFunction() + ' (' + t.getEventType() + ')';
  });
  lignes.push('Déclencheurs : ' + (declencheurs.length ? declencheurs.join(', ') : 'AUCUN'));
  if (declencheurs.join(' ').indexOf('surEnvoiFormulaire') < 0) {
    lignes.push('  -> PROBLÈME : le déclencheur surEnvoiFormulaire manque.');
  }

  ['FORM_DECLARATION_ID', 'FORM_VALIDATION_ID'].forEach(function (cle) {
    try {
      FormApp.openById(props.getProperty(cle));
      lignes.push(cle + ' : OK');
    } catch (err) {
      lignes.push(cle + ' : PROBLÈME (' + err.message + ')');
    }
  });

  var joueursSansEmail = lireJoueurs_().filter(function (j) {
    return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(j.email);
  }).map(function (j) { return j.nom; });
  lignes.push('Joueurs sans email valide : ' + (joueursSansEmail.length ? joueursSansEmail.join(', ') : 'aucun'));

  var dest = emailOrganisateur_();
  var test = ui.alert('Diagnostic Prime League',
    lignes.join('\n') + '\n\nEnvoyer un email de test à ' + dest + ' ?', ui.ButtonSet.YES_NO);
  if (test !== ui.Button.YES) return;
  try {
    MailApp.sendEmail({
      name: NOM_EXPEDITEUR,
      to: dest,
      subject: 'Prime League : email de test',
      body: "Si tu lis ceci, l'envoi d'emails par le script fonctionne vers " + dest + '.'
    });
    ui.alert('Email de test envoyé à ' + dest + '. Vérifie la boîte de réception et les spams.');
  } catch (err) {
    ui.alert("Échec de l'envoi : " + err.message);
  }
}

/* ------------------------------------------------------------------ */
/* Outils                                                              */
/* ------------------------------------------------------------------ */

function emailOrganisateur_() {
  var onglet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_PARAMETRES);
  var email = onglet ? String(onglet.getRange('B2').getValue()).trim() : '';
  return email || Session.getEffectiveUser().getEmail();
}

function echapper_(texte) {
  return String(texte)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
