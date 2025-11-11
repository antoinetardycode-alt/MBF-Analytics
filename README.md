# 💰 MBF-Analytics
### Classe VBA de méthodes financières basées sur RapidAPI

---

## 🚀 Objectif du projet

**MBF Analytics** est un projet collaboratif visant à créer une **classe VBA** regroupant un ensemble de **méthodes financières fiables, simples et vérifiées**, s’appuyant sur des **données issues de RapidAPI**.

Le but est d’offrir une bibliothèque **facile à utiliser** et **robuste**, permettant aux utilisateurs d’Excel de :

- récupérer des **données financières externes** (actions, devises, indices, etc.),
- tout en restant **100 % VBA natif**, sans dépendances externes lourdes.

---

## ⚙️ Fonctionnalités principales

| Type de méthode | Exemple | Description |
|------------------|----------|--------------|
| Indicateurs financiers | `Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")` |  Télécharge les indicateurs classiques annuels et trimestriels |
| Historique des cotations | `Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")` | Télécharge l'historique récent des cotations |

Toutes les méthodes sont :
- 🔍 **Simples à utiliser** (appel direct depuis Excel, illustré d'un exemple complet)
- 🔍 **Robustes** (gestion d’erreurs et d’API incluse)
- 🔍 **Vérifiées par un tiers** avant validation
- 🔍 **Documentées** avec un lien vers la page officielle de l'API et un exemple fonctionnel

---

## 🚀 Exemple d’utilisation

```
Sub test_MBFanalytics()

    Dim m As mbfAnalytics
    Set m = New mbfAnalytics
    
    ' Inscrire sa propre clé
    m.initKey "XXXXXXX"
    
    ' Test Bloomberg Financial
    Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")
        
    ' Test Real Time Quotes1
    Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")
    
End Sub
```

<p align="center">© MBF Assas — 2025</p>
