# Architecture Plane'R Crew

## Organisation

Apps Script

↓

Services

↓

Core

↓

Modules

↓

Interface utilisateur

---

## Règles

- Les modules ne communiquent jamais directement entre eux.

- Les composants Core sont réutilisables.

- Google Sheets est uniquement une base de données.

- Toute nouvelle fonctionnalité doit être modulaire.

---

## Organisation des fichiers

Core_*

↓

Services

↓

Modules

↓

Pages HTML

↓

Google Sheets

---

## Développement

Chaque mission suit le cycle :

Conception

↓

Développement

↓

Tests

↓

Validation

↓

Documentation   