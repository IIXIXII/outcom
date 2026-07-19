# Third-party notices

## Codex CLI

Outcom peut être distribué avec le binaire **Codex CLI** publié par OpenAI.

- Projet source : <https://github.com/openai/codex>
- Version épinglée : `0.144.6`
- Tag source : `rust-v0.144.6`
- Licence : Apache License 2.0
- Copyright : Copyright 2025 OpenAI

Le script `tools/Prepare-CodexRuntime.ps1` télécharge la licence correspondant
exactement au tag épinglé, en vérifie l’empreinte SHA-256 et la place à côté du
binaire sous le nom `Codex-CLI-LICENSE.txt`. Ce fichier de licence doit rester
présent dans tout paquet qui redistribue Codex CLI.

Outcom et Codex CLI restent des œuvres distinctes, distribuées sous leurs
licences respectives. Le nom OpenAI est utilisé uniquement pour identifier
l’origine du composant tiers.
