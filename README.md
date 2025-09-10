Ce dépôt contient des macros VBA et du RibbonX pour customiser Microsoft Visio
dans l’entreprise **Autographe**.

## 📦 Contenu
- `src/` : modules VBA et XML du ruban
- `docs/` : guides de style et tutoriels
- `templates/` : modèle Visio `.vstm` préconfiguré

## 🚀 Usage
1. Ouvrir `Autographe_Template.vstm` dans Visio
2. Accéder au ruban **Autographe**
3. Utiliser les macros intégrées (alignement, duplication de page, renommage…)

## 👨‍💻 Conventions
- Prefixes : `M_`, `C_`, `U_`
- Gestion erreurs : `On Error GoTo EH` + `M_Utils.LogErr`
- Callbacks RibbonX toujours dans `M_RibbonCallbacks`

Voir [docs/Autographe_Visio_Guide.md](docs/Autographe_Visio_Guide.md) pour plus de détails.
