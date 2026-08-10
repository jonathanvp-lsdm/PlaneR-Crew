# Roadmap Plane'R Crew

## Version actuelle

3.0

---

## Phase 1 — Fondations

✅ Architecture modulaire

✅ Dashboard

✅ Inscriptions

✅ Bénévoles

✅ Affectations

✅ Documents

✅ Administration

🚧 Gestion des éditions

⬜ Campagne d'inscription

⬜ Formulaire public

⬜ Validation

⬜ Communication

---

## Phase 2

⬜ Automatisation

⬜ Statistiques

⬜ Historique avancé

⬜ Tableau de bord intelligent

---

## Phase 3

⬜ Portail bénévole

⬜ Check-in

⬜ Signature électronique

⬜ QR Code

⬜ Notifications

✅ Mission 31.4 terminée
Framework
Ajout de la gestion des boutons (buttons) dans Core_ModalManager.
Compatibilité conservée avec l'ancien système footer.
Éditions
Remplacement du confirm() natif par une modale Bootstrap.
Suppression logique finalisée.
Interdiction de supprimer l'édition active.
Correction du calcul des IDs (plus de réutilisation après suppression logique).
Architecture
Validation de l'approche Core_ModalManager → Core_Dialog.
Décision de migrer progressivement des chaînes action vers des callbacks onClick.