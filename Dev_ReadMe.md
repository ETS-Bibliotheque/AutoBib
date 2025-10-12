# Installer uv
uv est une boite a outils moderne qui remplace pip pour gerer les environnements et les dependances.
1. Ouvrir PowerShell.
2. Lancer l'installation :
   ```powershell
   irm https://astral.sh/uv/install.ps1 | iex
   ```
3. Fermer puis rouvrir PowerShell, et verifier que tout fonctionne :
   ```powershell
   uv --version
   ```

# Creer un "Virtual ENVironment" avec uv
Se placer avec la console a la racine du projet (commande `cd`), puis creer l'environnement :
```powershell
uv venv .venv
```
uv telecharge automatiquement une version de Python compatible si necessaire.

# Activer le VENV (sur Windows)
En restant avec la console a la racine du projet (pour desactiver le VENV : remplacer activate par deactivate) :

### Command prompt
```batch
.venv\Scripts\activate
```

### PowerShell
```powershell
.venv\Scripts\Activate.ps1
```

# Installer les dependances avec uv
Avec le VENV actif et toujours a la racine du projet :
```powershell
uv sync
```
`uv sync` lit `pyproject.toml`, installe toutes les dependances dans le VENV et genere automatiquement le fichier de verrouillage `uv.lock`. Plus besoin de `requirements.txt`.
Si vous venez d'ouvrir un shell et que le VENV contient deja les dependances, utilisez :
```powershell
uv sync --locked
```
Cela verifie que `pyproject.toml` et `uv.lock` sont coherents sans retoucher aux versions.

# Mettre a jour les dependances
1. Activer le VENV (`.venv\Scripts\Activate.ps1` dans PowerShell).
2. Mettre a jour toutes les dependances definies dans `pyproject.toml` :
   ```powershell
   uv sync --upgrade
   ```
   Cette commande met a jour `uv.lock` avec les versions les plus recentes compatibles et reinstalle les dependances.
3. Pour cibler une dependance specifique :
   ```powershell
   uv sync --upgrade-package nom_du_package
   ```
4. Valider les changements dans `pyproject.toml` et `uv.lock` avant de les partager (les deux fichiers sont lies).

# Creer un executable a partir du code source
En etant toujours avec la console a la racine du projet et avec le VENV actif :
```batch
pyinstaller --add-data "Logos;Logos" --add-data "GABARIT.docx;." --add-data "GABARITCOLLABS.docx;." --add-data "GABARIT.xlsm;." --add-data "GABARITCOLLABS.xlsm;." --add-data "INFO.xlsx;." --add-data "maj_gen_py.bat;." --add-data "LICENSE;." --icon=Logos\ETS_Logo.ico --noconsole AutoBibPlus.py
```
Pour les prochaines fois, il suffit d'utiliser le fichier de specifications qui a ete cree (modifiable si besoin) :
```batch
pyinstaller AutoBibPlus.spec
```
Le dossier contenant l'executable et toutes les dependances se trouve alors dans `AutoBib\dist\AutoBibPlus`.
