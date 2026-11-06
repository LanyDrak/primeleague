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
  COMMENTAIRE: 11, CODE: 12
};
var NB_COLONNES_MATCHS = 12;

/* ------------------------------------------------------------------ */
/* Menu                                                                */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Prime League')
    .addItem('Installer les formulaires', 'installer')
    .addItem('Mettre à jour la liste des joueurs', 'majListeJoueurs')
    .addItem('Afficher les liens des formulaires', 'afficherLiens')
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
    "Déclare le résultat d'un match 1v1 sur FC 27. Ton adversaire recevra un email pour le confirmer. " +
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

  // Déclencheurs : réponse à un formulaire, modification de la liste des joueurs
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('surEnvoiFormulaire').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('surModification').forSpreadsheet(ss).onEdit().create();

  majListeJoueurs();
  afficherLiens();
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

function valeur_(namedValues, question) {
  var v = namedValues[question];
  return v && v.length ? String(v[0]).trim() : '';
}

function etoiles_(texte) {
  return Number(String(texte).replace(',', '.'));
}

function traiterDeclaration_(nv) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var onglet = ss.getSheetByName(ONGLET_MATCHS);
  var joueurs = lireJoueurs_();
  var j1 = valeur_(nv, Q_JOUEUR);
  var j2 = valeur_(nv, Q_ADVERSAIRE);
  var e1 = etoiles_(valeur_(nv, Q_ETOILES_JOUEUR));
  var e2 = etoiles_(valeur_(nv, Q_ETOILES_ADVERSAIRE));
  var b1 = Number(valeur_(nv, Q_BUTS_JOUEUR));
  var b2 = Number(valeur_(nv, Q_BUTS_ADVERSAIRE));

  var numero = Math.max(onglet.getLastRow(), 1); // ligne 1 = en-têtes
  var id = 'M' + ('000' + numero).slice(-4);
  var code = id + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
  var maintenant = new Date();

  var adversaire = joueurs.filter(function (j) { return j.nom === j2; })[0];
  var statut = STATUT_ATTENTE;
  var commentaire = '';
  if (j1 === j2) {
    statut = STATUT_REJETE;
    commentaire = 'Rejeté automatiquement : le joueur et son adversaire sont identiques.';
  } else if (!adversaire || !adversaire.email) {
    commentaire = "Email de l'adversaire manquant dans l'onglet Joueurs : validation à faire par l'organisateur.";
  }

  onglet.appendRow([id, maintenant, j1, e1, b1, j2, e2, b2, statut, maintenant, commentaire, code]);

  if (statut !== STATUT_ATTENTE) return;

  var resume = j1 + ' ' + b1 + ' - ' + b2 + ' ' + j2 +
    ' (étoiles : ' + j1 + ' ' + e1 + ', ' + j2 + ' ' + e2 + ')';

  if (adversaire && adversaire.email) {
    var lien = lienValidation_(code);
    MailApp.sendEmail({
      name: NOM_EXPEDITEUR,
      to: adversaire.email,
      subject: 'Prime League : confirme ton match contre ' + j1,
      htmlBody:
        '<p>Salut ' + echapper_(j2) + ',</p>' +
        '<p>' + echapper_(j1) + ' a déclaré le résultat suivant :</p>' +
        '<p style="font-size:16px"><b>' + echapper_(resume) + '</b></p>' +
        '<p><a href="' + lien + '">Confirmer ou contester ce résultat</a></p>' +
        '<p>Le match comptera au classement dès que tu l\'auras confirmé.</p>'
    });
  } else {
    MailApp.sendEmail({
      name: NOM_EXPEDITEUR,
      to: emailOrganisateur_(),
      subject: 'Prime League : match ' + id + ' sans email adversaire',
      htmlBody: '<p>Le match ' + id + ' (' + echapper_(resume) + ") n'a pas pu être envoyé en validation, " +
        "car l'adversaire n'a pas d'email dans l'onglet Joueurs. Tu peux le valider à la main dans l'onglet Matchs.</p>"
    });
  }
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
