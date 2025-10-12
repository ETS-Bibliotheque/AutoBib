
import sys, os, ctypes, win32gui, subprocess, requests
from datetime import datetime
from pathlib import Path

# Utilisation de QT pour la création de l'Interface Homme-Machine (IHM ou HMI en anglais)
from PySide6.QtWidgets import QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget, QPlainTextEdit, QMessageBox, QToolBar, QFileDialog
from PySide6.QtGui import QFont, QIcon, QFontMetrics, QAction
from PySide6.QtCore import Qt, Slot

# Importations locales
from Include.Front import ExitBox, CustomTextEdit, LoadingDialog, AchievedMessageBox, ReconfigMessageBox, InfoAPI, Info, Timer

# Appliquer une feuille de style CSS pour le texte en couleur
text_style_warning = '"color: #D35230"'
text_style_parameter = '"color: #AA1C4F"'
text_style_question = '"color: #0C5E31"'

def check_create_config(console: QPlainTextEdit, response: str, keys: list = None, first_time: bool = False):
    """Assure l'existence du fichier de configuration pybliometrics avant de poursuivre.

    :param console: composant console utilise pour afficher les invites et messages.
    :param response: saisie brute retournee par l'utilisateur via la console.
    :param keys: paire cle API / insttoken deja enregistree.
    :param first_time: indique si l'application s'execute pour la premiere fois.
    :return: tuple (validation, liste des cles normalisees).
    """
    import configparser
    from Include.pybliometrics.utils.constants import CONFIG_FILE
    from Include.pybliometrics.utils.create_config import create_config

    # Read/create config file (with fixture for RTFD.io)
    config = configparser.ConfigParser()
    config.optionxform = str

    # Fichier de configuration existant?
    if not CONFIG_FILE.exists():
        # Clef et token API déjà fournis?
        if keys == None:
            # Affiche dans un premier temps le message de bienvenu (permet aussi de garder la couleur noire sur les textes et tableaux par défaut)
            if first_time:
                console.append("Bienvenue sur <b>AutoBib+</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques et de collaborations de l'ÉTS!")
                console.append("")
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal.</p>'.format(text_style_parameter))
            # Rien ne se passe si la zone de texte rentrée est vide
            if response == '':
                return False, None
            # Sépare la zone de texte rentrée avec la virgule comme séparateur
            elements = response.split(',')
            # Check le nombre d'éléments et si le deuxième élément n'est pas nul
            if len(elements) < 2 or elements[1] == '':
                console.append('<p style={}>! Manque la clef et/ou le token</p>'.format(text_style_warning))
                console.append('')
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal.</p>'.format(text_style_parameter))
                return False, None
            # Suppression des espaces avant et après les éléments
            elements = [element.strip() for element in elements]
            # Test de la clef et du token par envoie d'une requête (API Scopus / AuthorRetrival sur Mohamed Cheriet)
            try:
                # Envoi de la requête à l'API Scopus
                response = requests.get("https://api.elsevier.com/content/author/author_id/56216876600", params={"apiKey": elements[0], "insttoken": elements[1], "httpAccept": "application/json"})
            except requests.exceptions.RequestException as e:
                error_message = "Une erreur s'est produite!\n\nSi l'erreur persiste, veuillez contacter le service technique de  votre établissement.\n\nDétails:\n" + str(e)
                error_dialog = QMessageBox(QMessageBox.Critical, "Erreur", error_message, QMessageBox.Ok)
                error_dialog.exec()
            
            # Vérification du statut de la réponse
            if response.status_code != 200:
                console.append('<p style={}>! Clef et/ou Token pour les API invalides (statut de la requête : {})</p>'.format(text_style_warning, response.status_code))
                console.append('')
                console.append('<p style={}>☼ Veuillez entrer votre clef API ainsi que votre token (séparés respectivement par une virgule) pour Scopus et SciVal :</p>'.format(text_style_parameter))
                return False, None
            
            # On peut passer à l'étape suivante : rentrer un chemin d'accès valide à un répertoire pour enregistrer les rapports par défaut
            console.append("\n\n")
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex : C:\\Users\\Name\\Documents).</p>".format(text_style_parameter))
            return False, elements
        
        # Rien ne se passe si la zone de texte rentrée est vide
        if response == '':
            return False, keys
        
        # Supprimer les espaces de la zone de texte rentrée, en faire un vrai chemin d'accès et vérifier s'il existe en local sur la machine
        docs_path = Path(response.strip())
        if not docs_path.is_dir():
            console.append("<p style={}>! Chemin d'accès à un dossier non-valide</p>".format(text_style_warning))
            console.append('')
            console.append("<p style={}>☼ Veuillez entrer le chemin d'accès ENTIER du répertoire où vous voulez enregistrer les rapports par défault (ex : C:\\Users\\Name\\Documents).</p>".format(text_style_parameter))
            return False, keys
        
        # Suppression du fichier de configuration non-abouti puis création du fichier complet
        subprocess.run("del %userprofile%\\.config\\pybliometrics.cfg", shell=True)
        config = create_config(keys=[keys[0]], insttoken=keys[1], docs_path=response.strip())

        # Gestion de l'affichage de la zone de text (effacage puis affichage des diverses informations)
        console.setPlainText('')
        console.append(f"Le fichier de configuration a été créé avec succès au chemin d'accès : {CONFIG_FILE}.")
        console.append("Pour plus de détails, veuillez consulter https://pybliometrics.rtfd.io/en/stable/configuration.html.")
        console.append('')

    console.append("""Bienvenue sur <b>AutoBib+</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques et de collaborations de l'ÉTS!
                    <br><br>Les commandes suivantes pourraient vous aider :
                    <br>&nbsp;&nbsp;- <b>Touche &lt;Entrée/Enter&gt;</b> :&nbsp;&nbsp;permet de sélectionner les paramètres <b>par défaut</b>
                    <br>&nbsp;&nbsp;- <b>Séparateur de valeurs</b> :&nbsp;&nbsp;utilisez la virgule
                    <br><br>La barre d'outils en rouge peut être déplacée à l'aide de sa ligne de points à son extrémité.
                    <br><br><b>Tapez vos commandes dans la barre d'entrée de texte tout en bas de la page</b>, puis validez-les en appuyant sur la touche &lt;Entrée/Enter&gt; de votre clavier.<br><br>""")

    # console.append('<span style="color: #0C5E31">● Veuillez entrer le nom puis le prénom de la personne (nom, prénom).</span>')  # Affiche la première interrogation dans la console
    console.append('<span style="color: #0C5E31">● Quel type de document souhaitez-vous produire? </span>')  # Affiche la première interrogation dans la console
    console.append('<span style="color: Black">   1. Fiche bibliométrique  </span>')
    console.append('<span style="color: Black">   2. Rapport de collaboration  </span>')
    return True, keys



class ConsoleWindow(QMainWindow):
    """Fenetre principale d'AutoBib+ pilotant le deroulement conversationnel.

    Elle orchestre la logique de generation des rapports bibliometriques et de collaboration.
    """
    def __init__(self):
        """Initialise les widgets, les menus et les variables d'etat de la console."""
        super().__init__() # Permet de récupérer le constructeur de la classe mère: QMainWindow
        script_path = os.getcwd()

        self.setWindowTitle("AutoBib+ : Logiciel d'automatisation des rapports d'analyses bibliométriques et de collaborations de l'ÉTS") # Définie le nom de la fenêtre
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/ETS_Logo.png"))  # Définit le logo

        # Définie la zone de texte (console)
        self.console = CustomTextEdit(self) # Instanciation

        # Définie la zone d'entrée de commandes de l'utilisateur
        self.input_box = QLineEdit(self) # Instanciation
        self.input_box.setStyleSheet("background-color: #DEDEDE; color: black; font-family: Consolas; font-size: 14pt; border: 2px solid black;") # Définie le CSS
        self.input_box.returnPressed.connect(self.handle_input) # Touche ENTER connectée à la méthode handle_input de la même classe
        # self.input_box.returnPressed.connect(self.printOK) # Touche ENTER connectée à la méthode handle_input de la même classe

        # Définie la barre d'outils
        toolbar = QToolBar('Toolbar', self)
        # Définition du style QSS pour la barre d'outils
        toolbar.setStyleSheet("""
            QToolBar {
                border: 0px solid black;
            };

            background-color: #EF3635;
        """)
        self.addToolBar(Qt.TopToolBarArea, toolbar)

        # Ajout action retour
        actRedo = QAction(QIcon(script_path + "/Logos/Back.svg"), "&Retour", self)
        # actRedo = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Back.svg"), "&Retour", self)
        actRedo.setShortcut('Ctrl+Z')
        actRedo.triggered.connect(self.retour)
        toolbar.addAction(actRedo)

        # Ajout d'un séparateur
        toolbar.addSeparator()

        # Ajout action raz
        actReset = QAction(QIcon(script_path + "/Logos/Reset.svg"), "&Remise à zéro", self)
        # actReset = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Reset.svg"), "&Remise à zéro", self)
        actReset.setShortcut('Ctrl+R')
        actReset.triggered.connect(self.raz)
        toolbar.addAction(actReset)

        toolbar.addSeparator()

        # Ajout action feuille blanche
        actResetALL = QAction(QIcon(script_path + "/Logos/WhiteSheet.svg"), "&Page blanche", self)
        # actResetALL = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/WhiteSheet.svg"), "&Page blanche", self)
        actResetALL.setShortcut('Ctrl+N')
        actResetALL.triggered.connect(self.whitesheet)
        toolbar.addAction(actResetALL)
        
        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action reconfiguration
        actReconfig = QAction(QIcon(script_path + "/Logos/Config.svg"), "&Reconfiguration", self)
        # actReconfig = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Config.svg"), "&Reconfiguration", self)
        actReconfig.setShortcut('F1')
        actReconfig.triggered.connect(self.reconfig)
        toolbar.addAction(actReconfig)

        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action Infos API
        actAPI = QAction(QIcon(script_path + "/Logos/API.svg"), "&Infos API", self)
        # actAPI = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/API.svg"), "&Infos API", self)
        actAPI.setShortcut('F2')
        actAPI.triggered.connect(self.API)
        toolbar.addAction(actAPI)

        toolbar.addSeparator()
        toolbar.addSeparator()

        # Ajout action Informations
        actInfos = QAction(QIcon(script_path + "/Logos/Infos.svg"), "&Informations", self)
        # actInfos = QAction(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/Infos.svg"), "&Informations", self)
        actInfos.setShortcut('F12')
        actInfos.triggered.connect(self.infos)
        toolbar.addAction(actInfos)

        # Définir la LoadingBox
        self.loading_dialog = LoadingDialog()

        # Organisation de la fenêtre Vertical Box et on met les deux éléments à la suite dans le bon ordre!
        layout = QVBoxLayout()
        layout.addWidget(self.console)
        layout.addWidget(self.input_box)

        # Créé la widget et ajoute le layout principal
        container = QWidget()
        container.setLayout(layout)
        container.setStyleSheet("background-color: #DEDEDE")  # Couleur d'arrière-plan de la widget principale

        # Set le widget sur la fenêtre principale
        self.setCentralWidget(container)

        # Initialisation
        self.text_style_question = '"color: #0C5E31"'
        self.input_box.setPlaceholderText("Entrez votre texte ici...")
        self.input_box.setFocus()
        # Instanciation du Timer
        self.timer = Timer()
        # Tableau pour les valeurs sur les clefs API
        self.infos_API = {'CitationOverview': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'}, \
                          'AuthorRetrieval': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'},\
                          'AuthorSearch': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'},\
                          'AuthorLookup': {'Date': 'None', 'X-RateLimit-Limit': 'None', 'X-RateLimit-Remaining': 'None', 'X-RateLimit-Reset': 'None'}}

        # Exécuter le fichier .bat de refresh du genpy pour une recherche plus rapide et éviter de futurs bugs liés aux fichiers de ce répertoire
        subprocess.run(os.path.dirname(os.path.abspath(__file__)) + "/maj_gen_py.bat", shell=True)

        # Backend:
        self.first_time = True # Permet de savoir si c'est la première fois que l'on rentre dans la fonction _affichageQuestions
        self.classeur = None # Variable qui va contenir l'instance du classeur Excel
        self.keys_valid = None
        # Variable pour stocker l'état courant de la machine à états
        validation, _ = check_create_config(self.console, '', first_time=True)
        self.state = 0 if validation else -1

        # Frontend
        '''console.append('<span style="color: Black">   1. Fiche bibliométrique  </span>')
        console.append('<span style="color: Black">   2. Rapport de collaboration  </span>')'''
        self.df_doc_type_selected = [0,0]
        self.tableauQuestions = [
            '<span style={}>● Quel type de document souhaitez-vous produire? </span><br><span style="color: Black">   1. Fiche bibliométrique  </span><br><span style="color: Black">   2. Rapport de collaboration  </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez entrer le nom puis le prénom de la personne (nom, prénom).</span>'.format(self.text_style_question),
            "<span style={}>● Quel.s type.s de documents souhaitez-vous exclure de la liste? (Par défaut : aucun)[Entrez le ou les numéro.s de l'index]</span>".format(self.text_style_question),
            "<span style={}>● Quelle est la plage d'années que vous choisissez? (Par défaut : {}, {} (c-à-d : 3ans et 5ans))(période carrière ajoutée par défaut, si aucune 3ème valeur n'est spécifiée)</span>".format(self.text_style_question, 'NONE', 'NONE'),
            "<span style={}>● Quels sont les 2 types de publications que vous voulez mettre en valeur? (Par défaut : {}, {} (c-à-d : {}, {}))[Entrez les numéros de l'index]</span><br><span style={}>--Option de combinaison sous le format [n1; n2] : syntaxe permettant de combiner des types de publications sous le nom du 1er type (n1).</span>".format(self.text_style_question, "0", "1", "NONE", "NONE", '"color: #75163F"'),
            'NULL',
            'NULL',
            'NULL',
            'NULL',
            'NULL',
            'NULL',
            '<span style={}>● Vous souhaitez produire un rapport de collaboration entre deux entités A et B. Veuillez choisir l\'entité A dans cette liste : </span><br><span style="color: Black">   1. Chercheur ou groupe de chercheurs </span><br><span style="color: Black">   2. Établissement ou groupe d\'établissements </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez choisir l\'entité A dans cette liste : </span><br><span style="color: Black">   1. Liste des ORN </span><br><span style="color: Black">   2. Réseau UQ </span><br><span style="color: Black">   3. Réseau ETS </span><br><span style="color: Black">   4. Autres </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez choisir l\'entité A dans cette liste : </span><br><span style="color: Black">   1. Liste des professeurs de l\'ETS </span><br><span style="color: Black">   2. Autres </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez entrer l\'identifiant (ou la liste des identifiants) Scopus de l\'entité A. <br> Si l\'entité A est un chercheur, vous pouvez aussi saisir son nom. <br> Utilisez la virgule comme séparateur si plusieurs identifiants à rentrer (ID1, ID2, ...)</span>'.format(self.text_style_question),
            '<span style={}>● Veuillez choisir l\'entité B dans cette liste : </span><br><span style="color: Black">   1. Chercheur ou groupe de chercheurs </span><br><span style="color: Black">   2. Établissement  ou groupe d\'établissements </span><br><span style="color: Black">   3. Pays </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez choisir l\'entité B dans cette liste : </span><br><span style="color: Black">   1. Liste des ORN </span><br><span style="color: Black">   2. Réseau UQ </span><br><span style="color: Black">   3. Réseau ETS </span><br><span style="color: Black">   4. Autres </span>'.format(self.text_style_question),
            '<span style={}>● Veuillez entrer l\'identifiant (ou la liste des identifiants) Scopus de l\'entité B. <br> Si l\'entité B est un chercheur, vous pouvez aussi saisir son nom. <br> Utilisez la virgule comme séparateur si plusieurs identifiants à rentrer (ID1, ID2, ...)</span>'.format(self.text_style_question),
            "<span style={}>● Quelle est la plage d'années que vous choisissez? (Par défaut : {}, {})</span>".format(self.text_style_question, 'NONE', 'NONE'),
            'NULL',
            'NULL'
        ]

    def handle_input(self):
        """Traite la saisie courante et fait avancer la machine a etats."""
        # print('dans handle_input')
        # Stock la réponse de l'utilisateur, efface la zone d'entrée de texte et affiche la réponse sur la zone de texte
        self.response = self.input_box.text()
        self.input_box.clear()
        self.console.append('<span style="color: blue">{}</span>'.format("► " + self.response)) # Afficher l'entrée de l'utilisateur

        # Si l'état courant est à -1, la création du fichier de configuration se fait alors
        if self.state == -1:
            # Check fichier de configuration déjà créé ou non
            validation, keys = check_create_config(self.console, self.response, self.keys_valid)
            if not validation:
                self.keys_valid = keys
                return
            self.state = 0
            return
        
        from Include.Tools import homonyme, selection_homonyme, tab_graph_Collab, \
            selection_types_de_documents, donnees_documents_graph_citations, selection_plages_annees, tab_graph_citations, \
            tab_graph_publications, Excel_part1, Excel_part2, valeurs_encadre, selection_2_types_docs, tab_graph_SNIP, \
            collaborationExtract, getSelectedYears, load_ETS_profs, findFuzzyMatches, countAuthorsInCollab, \
            countInstitutionsInCollab, countEntityAuthorsInCollab, add_affiliation_ids_to_list, load_UQ, load_ORN, load_ETS,\
            get_country_in_english, get_country_in_french, get_country_for_request,saveInter,count_document_types,\
            findOthersEtsAffiliations, findCollabCountryAffiliations, getEntityProfile, Excel_collabs_ETS_pays \
            # Excel_autres_collabs
        from Include.pybliometrics.scopus.author_search import AuthorSearch
        import pandas as pd
        
        # Afficher le message de chargement
        self.loading_dialog.show()

        # Forcer le rafraîchissement de l'interface utilisateur
        QApplication.processEvents()

        # Obtient la largeur d'un caractère (qui est fixe car nous sommes en police monospace)
        font_metrics = QFontMetrics(QFont("Consolas", 11))
        width_char = font_metrics.averageCharWidth()

        # Test l'exécution du script de l'état courant et test la validité des transitions possibles sinon message d'erreur
        try:
            # Machine à états
            match self.state:
                # État initial : recherche d'une personne par son nom et son prénom
                # État initial : choix du type de rapport à produire
                case 0:
                    # self.console.append('<p style={}>● Quel type de document souhaitez-vous produire? </p>'.format(text_style_question))
                    if self.response == '1': 
                        self.state += 1
                        self._affichageQuestions(1)
                    elif self.response == '2': 
                        self._affichageQuestions(11)
                        self.state = 11
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un type de document dans la liste proposée</p>'.format(text_style_warning))
                        # Fermer le message de chargement
                        self.loading_dialog.close()
                        return
                case 1:
                    # Validation du format de la requête de l'utilisateur
                    if ',' in self.response and not self.response.split(",")[1] =='':
                        last_name = self.response.split(",")[0]
                        first_name = self.response.split(",")[1]
                    else:
                        self.console.append('<p style={}>! Manque du séparateur (virgule) et/ou du prénom du personne</p>'.format(text_style_warning))
                        self.console.append('')
                        self.console.append('<p style={}>● Veuillez entrer le nom puis le prénom de la personne (nom, prénom).</p>'.format(text_style_question))
                        # Fermer le message de chargement
                        self.loading_dialog.close()
                        return                    
                    # Lancement du timer
                    self.timer.start()

                    # Recherche de la personne
                    self.search = AuthorSearch('AUTHLAST(' + last_name + ') and AUTHFIRST(' + first_name + ')', refresh=True)
                    self.infos_API['AuthorSearch'].update({key: self.search._header[key] for key in self.search._header if key in self.infos_API['AuthorSearch']})
                    
                    # Incrémentation en fonction du nombre d'homonyme
                    self.state += homonyme(self.search, self.console, int(self.width()/width_char)-10)
                    # self.state += 2                

                    # S'il n'y a pas d'homonymes
                    self._rechercheSurChercheur() if self.state == 3 else None

                # État 1 : cas où la recherche a mené à des homonymes, il faut alors choisir l'un d'entre eux      
                case 2:
                    # Vérifie si le ou les numéros d'index sont correctes
                    if selection_homonyme(self.response, self.search, self.console):
                        self.state += 1
                        self._rechercheSurChercheur(choix=int(self.response))

                # État 2 : certains documents doivent être exclus et export des premières données vers le doc Excel
                case 3:
                    # Séparer les types de documents sélectionnés par l'utilisateur
                    selected_types = self.response.split(',')
                    selected_types = [element.strip() for element in selected_types]

                    # Vérifie la conformité de la commande de l'utilisateur
                    if len(selected_types) == 1 and selected_types[0] == "":
                        self.index_list = self.df_doc_type.index.tolist()
                    elif selection_types_de_documents(selected_types, len(self.df_doc_type), self.console):
                        # Obtenir une liste d'entier puis mettre à jour la liste
                        selected_types = [int(x) for x in selected_types]
                        self.index_list = [x for x in self.df_doc_type.index.tolist() if x not in selected_types]
                    else:
                        self._affichageQuestions(2)
                        return
                    
                    self.state += 1

                    # Calcul, mise en forme des données pour le graphique des citations
                    self.docs_list, self.selected_eids_list, self.years = donnees_documents_graph_citations(self.au_retrieval, self.index_list, self.df_doc_type, self.console)
                    self.df, self.nom_prenom, header = tab_graph_citations(self.au_retrieval, self.selected_eids_list, self.docs_list, self.console, int(self.width()/width_char)-10)
                    
                    # Mise à jour du dictionnaire des données sur les API
                    self.infos_API['CitationOverview'].update({key: header[key] for key in header if key in self.infos_API['CitationOverview']})
                    
                    # Calcul, mise en forme des données pour les valeurs de l'encadré
                    self.en_tete, self.annee_10y_adapt, header = valeurs_encadre(self.authorEID, self.years)
                    
                    # Mise à jour du dictionnaire des données sur les API
                    self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                    self.excel, self.classeur = Excel_part1(self.df, self.nom_prenom, self.en_tete, self.annee_10y_adapt)

                    self._affichageQuestions(3)

                # État 3 : les plages d'années des histogrammes sont choisies ici
                case 4:
                    validation, self.years_list, self.df_doc_type_selected = selection_plages_annees(self.response, self.years, self.index_list, self.df_doc_type, self.console)
                    if validation:
                        self.state += 1

                        self._affichageQuestions(4)
                    else:
                        self._affichageQuestions(3)

                # État 4 : les 2 types de publications mis en avant sur le graphique des Pubications se fait ici AINSI que l'envoie de toutes les autres données pour
                # le doc Excel ainsi que l'appel des routines VBA pour réaliser la mise en forme des données et la création de la fiche bibliométrique Word
                case 5:
                    validation, self.type_list = selection_2_types_docs(self.response, self.df_doc_type_selected, self.console)
                    if validation:
                        self.state = 1
                                            
                        self.df_pub = tab_graph_publications(self.au_retrieval, self.selected_eids_list, self.years_list, self.type_list, self.console, int(self.width()/width_char)-10)

                        self.years_list = [[int(item) for item in sublist] for sublist in self.years_list]
                        self.df_SNIP, header = tab_graph_SNIP(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.df_Collab, header = tab_graph_Collab(console=self.console, author_id=self.authorEID, years_list=self.years_list, window_width=int(self.width()/width_char)-10)
                        self.infos_API['AuthorLookup'].update({key: header[key] for key in header if key in self.infos_API['AuthorLookup']})

                        nom_classeur = self.classeur.Name
                        nom_classeur = nom_classeur.split('.')[0]
                        Excel_part2(self.excel, self.classeur, self.df_pub, self.df_SNIP, self.df_Collab)
                        self.classeur = None

                        # Fermer le message de chargement
                        self.loading_dialog.close()

                        self.timer.stop()
                        achieved_msg = AchievedMessageBox(time=self.timer.get_elapsed_time())
                        achieved_msg.exec()
                        self.timer.reset()

                        self.console.append("\n")
                        self.console.append("<b>Rapport d'analyse bibliométrique créé avec succès!</b>")
                        self.console.append("\n\n")

                        # Mettre la fenêtre Word en premier plan
                        try:
                            win32gui.SetForegroundWindow(win32gui.FindWindow(None, nom_classeur + " - Word"))
                        except:
                            win32gui.SetForegroundWindow(win32gui.FindWindow(None, nom_classeur + '.docx' + " - Word"))
                        
                        self._affichageQuestions(1) #####
                    else:
                        self._affichageQuestions(4)
                case 11:
                    import configparser
                    from Include.pybliometrics.utils.constants import CONFIG_FILE
                    # Lancement du timer
                    self.timer.start()
                    self.fileNamePartA = ''
                    self.fileNamePartB = ''
                    self.indexAuteur = 0
                    if self.response == '1':
                        self.entiteA = self.response
                        self.state += 2 
                        self._affichageQuestions(13)
                    elif self.response == '2':
                        self.entiteA = self.response
                        self.state += 1 
                        self._affichageQuestions(12)
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un ensemble dans la liste proposée</p>'.format(text_style_warning))
                        return
                    config = configparser.ConfigParser()
                    config.read(CONFIG_FILE)

                    self.Keys = [config['Authentication']['APIKey'], config['Authentication']['InstToken']]
                case 12: 
                    self.listEntityA = []
                    self.reseauETS =  False
                    if self.response == '1' or self.response == '2' or self.response == '3':
                        if self.response == '1': 
                            listAffil = load_ORN(self.console)
                            self.fileNamePartA = 'Reseau_ORN'
                        elif self.response == '2': 
                            listAffil = load_UQ(self.console)
                            self.fileNamePartA = 'Reseau_UQ'
                        elif self.response == '3': 
                            listAffil = load_ETS(self.console)
                            self.fileNamePartA = 'Reseau_ETS'
                            self.reseauETS =  True
                        self.listEntityA = add_affiliation_ids_to_list(listAffil, self.listEntityA, self.console)
                        self.state += 3
                        self._affichageQuestions(15)
                    elif self.response == '4':
                        self.state += 2 
                        self._affichageQuestions(14)
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un ensemble dans la liste proposée</p>'.format(text_style_warning))
                        return
                case 13: 
                    self.listEntityA = []
                    if self.response == '1':
                        self.fileNamePartA = 'Profs_ETS'
                        listAffil = load_ETS_profs(self.console)
                        self.listEntityA = add_affiliation_ids_to_list(listAffil, self.listEntityA, self.console)
                        self.state += 2
                        self._affichageQuestions(15)
                    elif self.response == '2':
                        self.state += 1 
                        self._affichageQuestions(14)
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un ensemble dans la liste proposée</p>'.format(text_style_warning))
                        return

                case 14: 
                    self.listEntityA = []
                    # Verification de l'existence de l'entité
                    RechercheParId=False
                    if self.response == '' :
                        self._affichageQuestions(14)
                    elif ',' in self.response:
                        self.response = self.response.replace(" ", "")
                        self.response = self.response.replace("\n", "")
                        entities = self.response.split(',')
                        if self.entiteA == '1':
                            for caractere in entities:
                                if caractere.isdigit():
                                    RechercheParId = True
                            if RechercheParId is False:
                                results = getEntityProfile(self.entiteA, self.response, self.Keys, RechercheParId)
                                if results == 'NONE':
                                    self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                    self._affichageQuestions(14)
                                else:
                                    if len(results) == 1 : 
                                        self.console.append(f"Resumé du profil de l'auteur : \n")
                                        self.console.append(results[0])
                                        self.state += 1 
                                        self._affichageQuestions(15)
                                        for line in results[0].splitlines():
                                            if line.startswith("ID :"):
                                                identifiant = line.split(":")[1].strip()
                                                identifiant = identifiant.split("-")[2].strip()
                                                self.listEntityA.append(identifiant)
                                            if line.startswith("Nom :"):
                                                self.fileNamePartA = line.split(":")[1].strip()
                                                self.fileNamePartA = self.fileNamePartA.replace(" ", "_")
                                    else: 
                                        for index, result in enumerate(results): 
                                            self.console.append(f"Resumé du profil de l'auteur : ")
                                            self.console.append(f"Index : {index}")
                                            self.console.append(result)
                                            self.console.append("\n")
                                        self.console.append('<p style={}>● Veuillez choisir l\'index de l\'auteur </p>'.format(text_style_question))
                                        self.state = 19
                                        self.saveresults = results
                            else:
                                for index, entity in enumerate(entities):
                                    results = getEntityProfile(self.entiteA, entity, self.Keys, RechercheParId)
                                    if results == 'NONE':
                                        self.console.append(f"Auteur {index + 1}: \n")
                                        self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                    else:
                                        self.console.append(f"Resumé du profil de l'auteur {index + 1}: \n")
                                        self.console.append(results[0])
                                        self.listEntityA.append(entity)
                                self.fileNamePartA = 'Gr_Auteurs_A'
                                self.state += 1 
                                self._affichageQuestions(15)
                        elif self.entiteA == '2':
                            for index, entity in enumerate(entities):
                                results = getEntityProfile(self.entiteA, entity, self.Keys, RechercheParId)
                                if results == 'NONE':
                                    self.console.append(f"Institution {index + 1}: \n")
                                    self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                else:
                                    self.console.append(f"Resumé du profil de l'institution {index + 1}: \n")
                                    self.console.append(results)
                                    self.listEntityA.append(entity)
                            self.fileNamePartA = 'Gr_Institutions_A'
                            self.state += 1 
                            self._affichageQuestions(15)
                    else :
                        self.response = self.response.replace(" ", "")
                        self.response = self.response.replace("\n", "")
                        results = getEntityProfile(self.entiteA, self.response, self.Keys, True)
                        if results == 'NONE':
                            self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                            self._affichageQuestions(14)
                        else:
                            if self.entiteA == '1':
                                self.console.append("Resumé du profil de l'auteur : \n")
                                for line in results.splitlines():
                                    if line.startswith("Nom :"):
                                        self.fileNamePartA = line.split(":")[1].strip()
                                        self.fileNamePartA = self.fileNamePartA.replace(" ", "_")
                            elif self.entiteA == '2':
                                self.console.append("Resumé du profil de l'institution : \n")
                                self.fileNamePartA = 'Institution_A'
                            self.console.append(results)
                            self.listEntityA.append(self.response)
                            self.state += 1 
                            self._affichageQuestions(15)
                case 15:
                    if self.response == '2' : 
                        self.entiteB = self.response
                        self.state += 1
                        self._affichageQuestions(16)
                    elif self.response == '1' or self.response == '3':
                        self.entiteB = self.response
                        self.state += 2
                        self._affichageQuestions(17)
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un ensemble dans la liste proposée</p>'.format(text_style_warning))
                        return
                    
                case 16: 
                    self.listEntityB = []
                    if self.response == '1' or self.response == '2' or self.response == '3':
                        if self.response == '1': 
                            listAffil = load_ORN(self.console)
                            self.fileNamePartB = 'Reseau_ORN'
                        elif self.response == '2': 
                            listAffil = load_UQ(self.console)
                            self.fileNamePartB = 'Reseau_UQ'
                        elif self.response == '3': 
                            listAffil = load_ETS(self.console)
                            self.fileNamePartB = 'Reseau_ETS'
                        self.listEntityB = add_affiliation_ids_to_list(listAffil, self.listEntityB, self.console)
                        self.state += 2
                        self._affichageQuestions(18)
                    elif self.response == '4':
                        self.state += 1 
                        self._affichageQuestions(17)
                    else :
                        self.console.append('<p style={}>! Veuillez choisir un ensemble dans la liste proposée</p>'.format(text_style_warning))
                        return
                case 17:
                    self.listEntityB = []
                    RechercheParId = False
                    # Verification de l'existence de l'entité
                    if self.response == '' :
                        self._affichageQuestions(17)
                    elif ',' in self.response: 
                        self.response = self.response.replace(" ", "")
                        self.response = self.response.replace("\n", "")
                        entities = self.response.split(',')
                        if self.entiteB == '1':
                            for caractere in entities:
                                caractere = caractere.strip()
                                if caractere.isdigit():
                                    RechercheParId = True
                            if RechercheParId is False:
                                results = getEntityProfile(self.entiteB, self.response, self.Keys, RechercheParId)
                                if results == 'NONE':
                                    self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                    self._affichageQuestions(17)
                                else:
                                    if len(results) == 1 : 
                                        self.console.append(f"Resumé du profil de l'auteur : \n")
                                        self.console.append(results[0])
                                        self.state += 1 
                                        self._affichageQuestions(18)
                                        for line in results[0].splitlines():
                                            if line.startswith("ID :"):
                                                identifiant = line.split(":")[1].strip()
                                                identifiant = identifiant.split("-")[2].strip()
                                                self.listEntityB.append(identifiant)
                                            if line.startswith("Nom :"):
                                                self.fileNamePartB = line.split(":")[1].strip()
                                                self.fileNamePartB = self.fileNamePartB.replace(" ", "_")
                                    else: 
                                        for index, result in enumerate(results): 
                                            self.console.append(f"Resumé du profil de l'auteur : ")
                                            self.console.append(f"Index : {index}")
                                            self.console.append(result)
                                            self.console.append("\n")
                                        self.console.append('<p style={}>● Veuillez choisir l\'index de l\'auteur </p>'.format(text_style_question))
                                        self.state = 20
                                        self.saveresults = results

                            else:
                                for index, entity in enumerate(entities):
                                    results = getEntityProfile(self.entiteB, entity, self.Keys, RechercheParId)
                                    if results == 'NONE':
                                        self.console.append(f"Auteur {index + 1}: \n")
                                        self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                    else:
                                        self.console.append(f"Resumé du profil de l'auteur {index + 1}: \n")
                                        self.console.append(results[0])
                                        self.listEntityB.append(entity)
                                self.fileNamePartB = 'Gr_Auteurs_B'
                                self.state += 1 
                                self._affichageQuestions(18)
                        elif self.entiteB == '2':
                            for index, entity in enumerate(entities):
                                results = getEntityProfile(self.entiteB, entity, self.Keys, RechercheParId)
                                if results == 'NONE':
                                    self.console.append(f"Institution {index + 1}: \n")
                                    self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                else:
                                    self.console.append(f"Resumé du profil de l'institution {index + 1}: \n")
                                    self.console.append(results)
                                    self.listEntityB.append(entity)
                            self.fileNamePartB = 'Gr_Institutions_B'
                            self.state += 1 
                            self._affichageQuestions(18)
                    else :
                        self.response = self.response.replace(" ", "")
                        self.response = self.response.replace("\n", "")
                        if self.entiteB == '1' or self.entiteB == '2':
                            results = getEntityProfile(self.entiteB, self.response, self.Keys, True)
                            if results == 'NONE':
                                self.console.append('<p style={}>! Aucun profil trouvé </p>'.format(text_style_warning))
                                self._affichageQuestions(17)
                            else : 
                                self.listEntityB.append(self.response)
                                if self.entiteB == '1':
                                    self.console.append("Resumé du profil de l'auteur : \n")
                                    for line in results.splitlines():
                                        if line.startswith("Nom :"):
                                            self.fileNamePartB = line.split(":")[1].strip()
                                            self.fileNamePartB = self.fileNamePartB.replace(" ", "_")
                                elif self.entiteB == '2':
                                    self.console.append("Resumé du profil de l'institution : \n")
                                    self.fileNamePartB = 'Institution_B'
                                self.console.append(results)
                                self.state += 1 
                                self._affichageQuestions(18)

                        elif self.entiteB == '3':
                            self.country = self.response.strip()
                            self.country_for_request = get_country_for_request(self.country)
                            self.country_in_english = get_country_in_english(self.country)
                            self.fileNamePartB = get_country_in_french(self.country).replace(" ", "_")
                            if self.entiteB == '3' and self.country_in_english == 'NULL':
                                self.console.append('<p style={}>!Aucun pays ne correspond à cette saisie </p>'.format(text_style_warning))
                            else :
                                self.state += 1 
                                self._affichageQuestions(18)

                case 18:
                    self.start_year, self.end_year = getSelectedYears (self.response)
                    # Vérifie la conformité de la commande de l'utilisateur
                    if self.start_year == 'NULL' and self.end_year == 'NULL':
                        self.console.append('<p style={}>! Veuillez saisir une plage correcte (Format : début, fin)</p>'.format(text_style_warning))
                    else:
                        if self.entiteA == '2':
                            if len(self.listEntityA) == 1 and self.listEntityA[0] == '60026786':
                                self.fileNamePartA = 'ets'
                        
                        dateAjourdhui = str(datetime.now()).split(" ")[0]
                        filename = f'{dateAjourdhui}_collabs_{self.fileNamePartA}_{self.fileNamePartB}_{self.start_year}_{self.end_year}.xlsm'
                        #------------Recherche de données sur les collaborations entre l'entitéA et l'entitéB----------------------
                        if self.entiteA == '1' and self.entiteB == '1':
                            dfAllResult = collaborationExtract(researchersA= self.listEntityA, researchersB= self.listEntityB,\
                                                                start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                dfAllResult = count_document_types(dfAllResult)
                                df_authors_collab = countAuthorsInCollab(dfAllResult, self.Keys)
                                saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteurs=df_authors_collab)
                        elif self.entiteA == '1' and self.entiteB == '2':
                            dfAllResult = collaborationExtract(researchersA= self.listEntityA, institutionsB= self.listEntityB, \
                                                               start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                dfAllResult = count_document_types(dfAllResult)
                                df_authors_collab = countAuthorsInCollab(dfAllResult, self.Keys)
                                df_authors_entityB = countEntityAuthorsInCollab(dfAllResult, self.listEntityB, self.Keys)
                                saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteurs=df_authors_collab, dfAuteursB=df_authors_entityB)
                        elif self.entiteA == '2' and self.entiteB == '1':
                            dfAllResult = collaborationExtract(institutionsA= self.listEntityA, researchersB= self.listEntityB,\
                                                                start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                dfAllResult = count_document_types(dfAllResult)
                                df_authors_collab = countAuthorsInCollab(dfAllResult, self.Keys)
                                df_authors_entityA = countEntityAuthorsInCollab(dfAllResult, self.listEntityA, self.Keys)
                                saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteurs=df_authors_collab, dfAuteursA=df_authors_entityA)
            
                        elif self.entiteA == '2' and self.entiteB == '2':
                            dfAllResult = collaborationExtract(institutionsA= self.listEntityA, institutionsB= self.listEntityB,\
                                                                start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                # if (len(self.listEntityA) == 1 and self.listEntityA[0] == '60026786' ) or self.reseauETS is True: # ETS ou reseau ETS
                                #     df_prof_ets = load_ETS_profs(self.console)
                                #     matches_df, non_matches_df, fuzzy_matches = findFuzzyMatches(df_authors_collab, df_prof_ets, self.console)
                                #     other_ets_authors_df = findOthersEtsAffiliations(non_matches_df, dfAllResult)
                                #     other_authors_df = countEntityAuthorsInCollab(dfAllResult, self.listEntityB, self.Keys)
                                #     Excel_collabs_ETS_pays(filename, matches_df, other_ets_authors_df, other_authors_df, df_institutions, dfAllResult, fuzzy_matches, self.fileNamePartB, self.start_year, self.end_year, dateAjourdhui)
                                # else : 
                                    dfAllResult = count_document_types(dfAllResult)
                                    df_authors_entityA = countEntityAuthorsInCollab(dfAllResult, self.listEntityA, self.Keys)
                                    df_authors_entityB = countEntityAuthorsInCollab(dfAllResult, self.listEntityB, self.Keys)
                                    saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteursA=df_authors_entityA, dfAuteursB=df_authors_entityB)
                        elif self.entiteA == '1' and self.entiteB == '3':
                            dfAllResult = collaborationExtract(researchersA= self.listEntityA, country=self.country_for_request,\
                                                                start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                dfAllResult = count_document_types(dfAllResult)
                                df_institutions = countInstitutionsInCollab(dfAllResult, collabCountry=self.country_in_english)
                                df_authors_collab = countAuthorsInCollab(dfAllResult, self.Keys)
                                df_authors_entityB = findCollabCountryAffiliations(df_authors_collab, dfAllResult, self.country_in_english, self.Keys)
                                saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteurs=df_authors_collab, dfAuteursB=df_authors_entityB, dfInstitutions=df_institutions)
                        elif self.entiteA == '2' and self.entiteB == '3':
                            dfAllResult = collaborationExtract(institutionsA= self.listEntityA, country=self.country_for_request,\
                                                                start_year=self.start_year, end_year=self.end_year, keys = self.Keys, console=self.console)
                            if dfAllResult is None:
                                self.state = 0
                                self._affichageQuestions(0)
                            else:
                                dfAllResult = count_document_types(dfAllResult)
                                df_authors_collab = countAuthorsInCollab(dfAllResult, self.Keys)
                                df_institutions = countInstitutionsInCollab(dfAllResult, collabCountry=self.country_in_english)
                                if (len(self.listEntityA) == 1 and self.listEntityA[0] == '60026786' ) or self.reseauETS is True: # ETS ou reseau ETS
                                    df_prof_ets = load_ETS_profs(self.console)
                                    matches_df, non_matches_df, fuzzy_matches = findFuzzyMatches(df_authors_collab, df_prof_ets, self.console)
                                    other_ets_authors_df = findOthersEtsAffiliations(non_matches_df, dfAllResult)
                                    other_authors_df = findCollabCountryAffiliations(non_matches_df, dfAllResult, self.country_in_english, self.Keys)
                                    # filename = f'collaborations_ets_{self.fileNamePartB}_{self.start_year}_{self.end_year}.xlsm'
                                    Excel_collabs_ETS_pays(filename, matches_df, other_ets_authors_df, other_authors_df, df_institutions, dfAllResult, fuzzy_matches, self.fileNamePartB, self.start_year, self.end_year, dateAjourdhui)
                                else : 
                                    df_authors_entityA = countEntityAuthorsInCollab(dfAllResult, self.listEntityA, self.Keys)
                                    df_authors_entityB = findCollabCountryAffiliations(df_authors_collab, dfAllResult, self.country_in_english, self.Keys)
                                    df_institutions = countInstitutionsInCollab(dfAllResult, collabCountry=self.country_in_english)
                                    saveInter(fileName=filename, dfAllResults=dfAllResult , dfAuteursA=df_authors_entityA, dfAuteursB=df_authors_entityB, dfInstitutions=df_institutions)

                       # Fermer le message de chargement
                        self.loading_dialog.close()

                        self.timer.stop()
                        if dfAllResult is not None:
                            achieved_msg = AchievedMessageBox(time=self.timer.get_elapsed_time())
                            achieved_msg.exec()
                        self.timer.reset()
                        if dfAllResult is not None:
                            # dfAllResult.to_excel(filename, index=False)
                            self.console.append("\n")
                            self.console.append("<b>Rapport de collaboration créé avec succès!</b>")
                            self.console.append("\n\n")
                            self.state = 0
                            self._affichageQuestions(0)
                # Cette partie a été rajoutée au dernier moment. D'ou son emplacement. 
                # Ramener ces etapes apres l'etape 14 et mettre à jour le code en consequence
                case 19: 
                    index = int(self.response)
                    if index >= len(self.saveresults) : 
                        self.console.append('<p style={}>! Veuillez choisir un index dans la liste proposée</p>'.format(text_style_warning))
                    else :
                        line =  self.saveresults[index].splitlines()
                        identifiant = line[1].split(":")[1].strip()
                        identifiant = identifiant.split("-")[2].strip()
                        self.listEntityA.append(identifiant)
                        self.fileNamePartA = line[0].split(":")[1].strip()
                        self.fileNamePartA = self.fileNamePartA.replace(" ", "_")
                        self.state = 15
                        self._affichageQuestions(15)
                case 20: 
                    index = int(self.response)
                    if index >= len(self.saveresults) : 
                        self.console.append('<p style={}>! Veuillez choisir un index dans la liste proposée</p>'.format(text_style_warning))
                    else :
                        line =  self.saveresults[index].splitlines()
                        identifiant = line[1].split(":")[1].strip()
                        identifiant = identifiant.split("-")[2].strip()
                        self.listEntityB.append(identifiant)
                        self.fileNamePartB = line[0].split(":")[1].strip()
                        self.fileNamePartB = self.fileNamePartB.replace(" ", "_")
                        self.state = 18
                        self._affichageQuestions(18)
                    #--------------------------------------------------------------------------------------------------------------------------------


        # Une erreur est rencontrée lors de l'exécution d'un des états, alors une boîte de dialogue s'affiche avec les détails de l'erreur
        except Exception as e:
            error_message = "Une erreur s'est produite!\n\nSi l'erreur persiste, veuillez contacter le service technique de  votre établissement.\n\nDétails :\n" + str(e)
            error_dialog = QMessageBox(QMessageBox.Critical, "Erreur", error_message, QMessageBox.Ok)
            error_dialog.exec()
            print(e)
            self.state = 0
            self._affichageQuestions(self.state)

        # Déplacer le QTextEdit à sa toute fin
        self.console.verticalScrollBar().setValue(self.console.verticalScrollBar().maximum())

        # Fermer le message de chargement
        self.loading_dialog.close()

    def _affichageQuestions(self, affichage_type: int):
        """Affiche la question ou le message correspondant a l'etat du deroulement."""
        if affichage_type == 3:
            # Supprimer le troisième élément en utilisant pop()
            self.tableauQuestions.pop(3)
            self.tableauQuestions.insert(3, "<span style={}>● Quelle est la plage d'années que vous choisissez? (Par défaut : {}, {} (c-à-d : 3ans et 5ans))(période carrière ajoutée par défaut, si aucune 3ème valeur n'est spécifiée)</span>".format(self.text_style_question, datetime.now().year-3, datetime.now().year-5))
        elif affichage_type == 4:
            # Supprimer le quatrième élément en utilisant pop()
            self.tableauQuestions.pop(4)
            liste_types_selec = self.df_doc_type_selected['Type de documents'].to_list()
            default_first = self.index_list[0]
            default_second = self.index_list[1] if len(self.index_list) > 1 else 'N/A'
            label_first = liste_types_selec[0]
            label_second = liste_types_selec[1] if len(liste_types_selec) > 1 else 'N/A'
            prompt = (
                "<span style={}>? Quels sont les 2 types de publications a mettre en avant ? (Par defaut : {}, {} soit {}, {}) [Entrez les numeros d'index]</span>"
                "<br><span style={}>--Option de combinaison [n1; n2] pour regrouper sous le premier type.</span>"
            ).format(self.text_style_question, default_first, default_second, label_first, label_second, '"color: #75163F"')
            self.tableauQuestions.insert(4, prompt)
        elif affichage_type == 18:
            self.tableauQuestions.pop(18)
            self.tableauQuestions.insert(18, "<span style={}>● Quelle est la plage d'années que vous choisissez? (Par défaut : {}, {})</span>".format(self.text_style_question, datetime.now().year-5, datetime.now().year))
        elif affichage_type == 17:
            if self.entiteB == '1' or self.entiteB == '2':
                self.tableauQuestions.pop(17)
                self.tableauQuestions.insert(17, "<span style={}>●   Veuillez entrer l\'identifiant (ou la liste des identifiants) Scopus de l\'entité B. <br> Si l\'entité B est un chercheur, vous pouvez aussi saisir son nom. <br> Utilisez la virgule comme séparateur si plusieurs identifiants à rentrer (ID1, ID2, ...) </span>".format(self.text_style_question))
            elif self.entiteB == '3':
                self.tableauQuestions.pop(17)
                self.tableauQuestions.insert(17, "<span style={}>●  Veuillez entrer le nom du pays </span>".format(self.text_style_question))
        self.console.append('')
        self.console.append(self.tableauQuestions[affichage_type])

    def _rechercheSurChercheur(self, choix: int = 0):
        """Interroge Scopus pour la personne selectionnee et actualise le cache local."""
        from Include.Tools import retrieval, tous_les_docs_chercheur

        self.authorEID, self.au_retrieval = retrieval(choix, self.search, self.console)
        self.infos_API['AuthorRetrieval'].update({key: self.au_retrieval._header[key] for key in self.au_retrieval._header if key in self.infos_API['AuthorRetrieval']})

        self.df_doc_type = tous_les_docs_chercheur(self.au_retrieval, self.console)

        self._affichageQuestions(2)


    def closeEvent(self, event):
        """Propose une confirmation avant de fermer la fenetre de l'application."""
        message_box = ExitBox(self) # Instanciation
        message_box.exec()

        if message_box.clickedButton() == message_box.buttonYes:
            event.accept()

            # Arrêt du chronomètre et remise à zéro
            self.timer.stop()
            self.timer.reset()

            # Fermer le gabarit sans l'enregistrer s'il est ouvert 
            self.classeur.Close(SaveChanges=False) if self.classeur is not None else None
        else:
            event.ignore()


    @Slot()
    def retour(self):
        """Revient a l'etape precedente lorsque le retour est autorise."""
        if self.state != 0 and self.state != -1:
            if (self.state > 1 and  self.state < 6) or (self.state > 11 and  self.state < 21):
                if self.state == 13 or self.state == 16:
                    self.state -= 2
                else:
                    if self.state == 19:
                        self.state = 14
                    elif self.state == 20:
                        self.state = 17
                    else:
                        self.state -= 1
        else: 
            self.state = 0
        self._affichageQuestions(self.state)
        

    @Slot()
    def raz(self):
        """Reinitialise l'interaction et renvoie vers l'accueil."""
        if self.state != 0 and self.state != -1:
            self.state = 0
            self._affichageQuestions(self.state)

            # Fermer le gabarit sans l'enregistrer s'il est ouvert
            if self.classeur is not None:
                self.classeur.Close(SaveChanges=False) 
                self.classeur = None
            
    @Slot()
    def reconfig(self):
        """Ouvre l'assistant de reconfiguration pour mettre a jour identifiants et chemins."""
        message_box = ReconfigMessageBox(self) # Instanciation
        message_box.exec()
        if self.state != -1 and message_box.clickedButton() == message_box.buttonYes:            
            self.state = -1
            self.console.setPlainText('')
            subprocess.run("del %userprofile%\\.config\\pybliometrics.cfg", shell=True)
            check_create_config(self.console, '', None, True)

    @Slot()
    def API(self):
        """Affiche les informations d'API memorisees pour la session en cours."""
        message_box = InfoAPI(self.infos_API) # Instanciation
        message_box.exec()

    @Slot()
    def infos(self):
        """Affiche la boite d'information generale sur AutoBib+."""
        message_box = Info() # Instanciation
        message_box.exec()

    @Slot()
    def whitesheet(self):
        """Vide la console sans modifier l'etat courant."""
        if self.state != -1:
            self.console.setPlainText('')
            
            self.console.append("""Bienvenue sur <b>AutoBib+</b>, le logiciel qui vous permez de générer automatiquement les rapports d'analyses bibliométriques et de collaborations de l'ÉTS!
                                    <br><br>Les commandes suivantes pourraient vous aider :
                                    <br>&nbsp;&nbsp;- <b>Touche &lt;Entrée/Enter&gt;</b> :&nbsp;&nbsp;permet de sélectionner les paramètres <b>par défaut</b>
                                    <br>&nbsp;&nbsp;- <b>Séparateur de valeurs</b> :&nbsp;&nbsp;utilisez la virgule
                                    <br><br>La barre d'outils en rouge peut être déplacée à l'aide de sa ligne de points à son extrémité.
                                    <br><br><b>Tapez vos commandes dans la barre d'entrée de texte tout en bas de la page</b>, puis validez-les en appuyant sur la touche &lt;Entrée/Enter&gt; de votre clavier.<br>""")

            self.state = 1
            self.raz()



# Main loop
if __name__ == "__main__":
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ETS.Automatisation_Rapports_Bibliometriques') # Permet que l'OS voit l'exécution du script indépendante de Python et donc de changer le logo lors du script

    app = QApplication(sys.argv) # Instanciation d'une application QT avec les arguments système
    app.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/Logos/ETS_Logo.png")) # Affiche l'icône

    window = ConsoleWindow() # Instanciation de la fenêtre que nous venons de créer
    window.setStyleSheet("background-color: white; border: NONE;") # Couleur d'arrière-plan de la fenêtre

    window.showMaximized() # Affiche en plein écran

    sys.exit(app.exec()) # Pour la fermeture de l'application
