# PrimeLeague

Classement FIFA interne Primelis : un classement ELO pour les matchs 1v1 sur FC 27.

## Fonctionnement

1. Un joueur déclare son match avec un Google Form (score et nombre d'étoiles de chaque équipe).
2. L'adversaire reçoit un email « PrimeLeague Validator » pour confirmer ou contester le résultat.
3. Seuls les matchs confirmés comptent. Le Google Sheet recalcule l'ELO match par match.
4. La page web `site/index.html` lit le classement publié depuis le Sheet et l'affiche.

## Règles de calcul

- ELO de départ : 1000, facteur K : 32.
- Écart de buts : points multipliés par 1 (1 but d'écart ou nul), 1,5 (2 buts), puis (11 + écart) / 8 au-delà.
- Équipe jouée : chaque étoile d'écart entre les deux équipes compte comme 50 points d'ELO dans le score attendu. Gagner avec l'équipe la plus faible rapporte donc plus.

Tous ces réglages se modifient dans l'onglet Paramètres du Google Sheet.

## Contenu

| Fichier | Rôle |
|---|---|
| `apps-script/Code.gs` | Script à coller dans le Google Sheet (Extensions > Apps Script). Crée les deux formulaires, envoie les emails de validation et met à jour la liste des joueurs. |
| `site/index.html` | Page web statique du classement. |
| `site/photos/` | Photos des joueurs au format 25/35, nommées avec le prénom en minuscules et sans accent (`karim.jpg`). |

## Brancher la page au Google Sheet

1. Dans le Sheet : Fichier > Partager > Publier sur le Web, onglet **Classement**, format **CSV**. Copier le lien.
2. Même chose pour l'onglet **Calcul ELO**.
3. Coller les deux liens dans `CONFIG.csvClassement` et `CONFIG.csvMatchs`, en haut du script de `site/index.html`.

Seuls ces deux onglets deviennent publics. Les emails de l'onglet Joueurs restent privés. Sans ces liens, la page affiche des données de démonstration.
