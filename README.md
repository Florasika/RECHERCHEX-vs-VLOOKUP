# 🔍 Jour 3 / 10 — RECHERCHEX vs VLOOKUP vs INDEX+EQUIV

> **Série : 10 Days of Excel** · Jour 3/10  
> Concepts : RECHERCHEX · VLOOKUP · INDEX · EQUIV · SIERREUR

---

## 📁 Structure du fichier

```
recherchex_jour3.xlsx
│
├── PRODUITS      ← 10 produits source (ID, Nom, Prix, Stock, Fournisseur...)
├── COMPARAISON   ← Les 3 formules côte à côte sur le même cas concret
└── EXERCICES     ← 5 exercices avec espace de réponse + solutions corrigées
```

---

## 🔑 Les 3 formules expliquées

### 1. RECHERCHEX — La nouvelle référence ✅
```excel
=RECHERCHEX(
    valeur_cherchée,        ← ce que tu cherches
    tableau_recherche,      ← où chercher (colonne des IDs)
    tableau_retour,         ← ce que tu veux récupérer
    "Non trouvé",           ← valeur si absent (OPTIONNEL)
    0,                      ← correspondance exacte
    1                       ← ordre de recherche
)
```
**Avantages :**
- Cherche dans les 2 sens (gauche ET droite)
- Valeur par défaut intégrée — plus besoin de SIERREUR
- Résistante aux insertions de colonnes
- Syntaxe lisible

**Limitation :** Excel 2021 / Microsoft 365 uniquement

---

### 2. VLOOKUP / RECHERCHEV — La méthode classique ⚠️
```excel
=RECHERCHEV(valeur_cherchée, tableau_complet, num_colonne, FAUX)

# Avec gestion d'erreur :
=SIERREUR(RECHERCHEV(valeur_cherchée, tableau_complet, 4, FAUX), "Non trouvé")
```
**Avantages :**
- Compatible Excel 2010, 2013, 2016, 2019
- Largement connue et documentée

**Limitations :**
- Cherche UNIQUEMENT vers la droite
- Le numéro de colonne est fragile — insérer une colonne = formule incorrecte
- Retourne `#N/A` sans message — nécessite SIERREUR en plus

---

### 3. INDEX + EQUIV — La formule pro 🔧
```excel
=INDEX(
    colonne_retour,                              ← colonne à retourner
    EQUIV(valeur_cherchée, colonne_recherche, 0) ← position de la valeur
)
```
**Comment ça fonctionne :**
- `EQUIV("PRD003", A2:A11, 0)` → retourne `3` (3ème ligne)
- `INDEX(D2:D11, 3)` → retourne la valeur ligne 3 de la colonne D

**Avantages :**
- Compatible toutes versions
- Cherche dans les 2 sens
- Très flexible — combine avec MIN, MAX pour trouver le moins cher etc.

---

## 📊 Tableau comparatif

| Critère | RECHERCHEX | VLOOKUP | INDEX+EQUIV |
|---------|-----------|---------|-------------|
| Excel 365/2021 | ✓ | ✓ | ✓ |
| Excel 2016/2019 | ✗ | ✓ | ✓ |
| Cherche à gauche | ✓ | ✗ | ✓ |
| Valeur par défaut | ✓ | ✗ | Avec SIERREUR |
| Lisibilité | ★★★ | ★★ | ★ |
| Résistance insertions | ★★★ | ★ | ★★★ |

---

## ✏️ Les 5 exercices

| # | Formule | Difficulté |
|---|---------|-----------|
| 1 | RECHERCHEX simple | ⭐ |
| 2 | RECHERCHEX + valeur par défaut | ⭐⭐ |
| 3 | INDEX+EQUIV sur le MIN | ⭐⭐⭐ |
| 4 | RECHERCHEX inversé (droite → gauche) | ⭐⭐ |
| 5 | SIERREUR + VLOOKUP | ⭐⭐ |

---

## 💡 Règle simple à retenir

```
Excel 365 / 2021  →  RECHERCHEX
Excel 2016 / 2019 →  INDEX + EQUIV
Jamais VLOOKUP seul (toujours entourer de SIERREUR)
```
