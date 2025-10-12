# © 2023 Benjamin Lepourtois <benjamin.lepourtois@gmail.com>
# © 2024 Adji Toure <adji.toure.dev@gmail.com>
# Copyright: All rights reserved.
# See the license attached to the root of the project.

"""
Projet d'Automatisation de Rapports d'Analyses Bibliométriques:

Ce programme fait partie intégrante du projet de conception et développement d'outils automatisés pour la réalisation de rapports d'analyses bibliométriques.

Contexte:

  ● Stage de 12 semaines sur l'été 2023 (12 juin au 1er septembre) dans l'École de Technologie Supérieure, Montréal, Canada
  ● Mission principale:
Développer des outils permettant l'automatisation de certaines étapes de production de rapports d'analyses bibliométriques 
destinés à aider les chercheurs et chercheuses dans la planification de la mesure de l'impact de leurs contributions scientifiques.

Approche choisie:

Nous avons choisi d'utiliser un script Python pour gérer toute l'automatisation des rapports.
  ● Extraction des données: par les API des différentes plateformes utilisées (Scopus et SciVal) à l'aide de la bibliothèque publique "pybliometrics"
  ● Traitement des données: en Python à l'aide de la bibliothèque "pandas"
  ● Interface Homme-Machine: en QT avec une interface très simpliste basée sur une boîte de dialogue
  ● Exportation des données: en Python à l'aide de la bibliothèque "pywin32" vers un fichier "Workbook" MacroExcel (.xlsm)
  ● Mise en forme Excel: avec des routines VBA appelées par le script Python
  ● Réalisation du rapport Word: avec des routines VBA, appelées par le script Python, qui exportent les données et les graphiques réalisés sur un document Word
"""

""" 
Projet repris par Adji Touré
Contexte :
  ● Stage de 12 semaines sur l'été 2024 (3 juin au 23 août) à l'École de Technologie Supérieure, Montréal, Canada.

Mission principale : 
  ● Continuer le développement de l'outil en intégrant la production de rapports de colloaborations entre deux entités.

Approche adoptée et tâches réalisées:
  ● Observation de la méthodologie existante : Analyse des processus et méthodes actuels utilisés pour la réalisation des rapports bibliométriques.
  ● Participation aux échanges avec les fournisseurs : Interaction avec les partenaires externes pour mieux comprendre les outils et services disponibles.
  ● Évaluation de différentes approches : Identification et test de la faisabilité de diverses approches pour le projet, analyse des avantages et inconvénients de chacune, et présentation de recommandations à l'équipe pour sélectionner l'approche optimale.
  ● Maîtrise des bibliothèques et APIs : Étude approfondie de la documentation sur l'exploitation des APIs, notamment celles d'Elsevier, et acquisition de compétences en utilisant des bibliothèques comme pybliometrics pour l'extraction de données et pandas pour leur traitement.
  ● Développement de scripts pour la détection de noms : Création d'un script Python pour détecter et identifier les noms des professeurs de l'établissement en utilisant des techniques de correspondance approximative.
  ● Automatisation des rapports : Développement de scripts Python pour :
        Extraction des données bibliométriques via l'API Scopus.
        Traitement et calcul des indicateurs requis.
        Exportation des données vers Excel et intégration des graphiques dans des rapports Word via des routines VBA.
        Amélioration de l'Interface Homme-Machine (IHM) : Développement d'une nouvelle branche pour l'IHM en Qt.
"""

import os, threading, time, pytz
from datetime import datetime

# Utilisation de QT pour la création de l'Interface Homme-Machine (IHM ou HMI en anglais)
from PySide6.QtWidgets import QMessageBox, QTextEdit, QDialog, QVBoxLayout, QLabel, QWidget, QPlainTextEdit, QTabWidget
from PySide6.QtGui import QFont, QIcon, QPixmap, QCursor
from PySide6.QtCore import Qt


class ExitBox(QMessageBox):
    """Boite modale confirmant la fermeture d'AutoBib+."""
    def __init__(self, parent=None):
        """Initialise la boite de confirmation avec le style applique a l'ensemble du logiciel."""
        super().__init__(parent) # Permet de récupérer le constructeur de la classe mère: QMainWindow
              
        # Définie la feuille de style pour les différents composants de l'ExitBox
        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 12pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)

        # Définie l'icone, le titre, et la question posée
        self.setIcon(QMessageBox.Question)
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setWindowTitle("Confirmation")
        self.setText("\nÊtes-vous sûr de vouloir quitter?")

        # Définie les boutons pour répondre à la question (role et mise en forme)
        self.buttonYes = self.addButton("Oui", QMessageBox.YesRole)
        self.buttonYes.setFont(QFont("Arial", 11, QFont.Bold))
        self.buttonNo = self.addButton("Non", QMessageBox.NoRole)
        self.buttonNo.setFont(QFont("Arial", 11, QFont.Bold))

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/ETS_Logo.png")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(64, 64))  # Redimensionner l'icône et l'assigner


class CustomTextEdit(QTextEdit):
    """Classe de la zone de texte (console), redéfinition de la classe QTextEdit pour répondre à nos besoins
    """
    def __init__(self, parent=None):
        """Configure la console en lecture seule utilisee pendant l'interaction."""
        super().__init__(parent)

        self.setStyleSheet("background-color: #DEDEDE; color: black; font-family: Consolas; font-size: 11pt; border: NONE;") # Définie le CSS de la console
        self.setReadOnly(True) # En lecture seule
        self.setLineWrapMode(QTextEdit.WidgetWidth)  # Mode de retour à la ligne en fonction de la largeur du widget   WidgetWidth


class LoadingDialog(QDialog):
    """Classe de la boîte de dialogue de chargement du programme
    """
    def __init__(self):
        """Cree la boite de chargement affichee pendant les traitements."""
        super().__init__()

        # Définie un arangement vertical pour la box avec un Label à l'intérieur avec le message d'attente
        layout = QVBoxLayout()
        self.label = QLabel("Exécution en cours, veuillez patienter...")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

        # Définie le titre et l'icone de la box
        self.setWindowTitle("Chargement...")
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setModal(True)

        # Afficher une icône de run pour le curseur de la souris
        self.setCursor(QCursor(Qt.WaitCursor))


class AchievedMessageBox(QMessageBox):
    """Classe du message de succès de la création du rapport
    """
    def __init__(self, time: int = 0) -> None:
        """Prepare la boite de succes qui resume le rapport genere."""
        super().__init__()

        # Définie la feuille de style des différents composants de la box
        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 10pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)

        # Définie l'icone, le titre mais aussi le texte du contenu de la Box
        from .pybliometrics.utils.startup import DOCS_PATH
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png")) 
        self.setWindowTitle("Rapport réalisé")
        self.setText(f"Félicitations!\n\nRapport d'analyse bibliométrique créé avec succès.\n\nLe rapport est enregistré par défaut dans le répertoire suivant :\n{DOCS_PATH[0]}\n\nCe rapport a été généré en : {int(time/60)}min {time%60}s")

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/ETS_Logo.png")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(64, 64))  # Redimensionner l'icône et l'assigner
    

class ReconfigMessageBox(QMessageBox):
    """Confirm that the user really wants to reset and reconfigure AutoBib+."""
    def __init__(self, parent=None):
        """Demande une confirmation avant de relancer la configuration."""
        """Initialise la boite de confirmation avec le style applique a l'ensemble du logiciel."""
        super().__init__(parent)
              
        # Définie la feuille de style des différents composants de la box
        self.setStyleSheet("""
            QMessageBox { background-color: #DEDEDE; font-size: 12pt; }
            QLabel { background-color: #DEDEDE; }
            QPushButton { background-color: #EF3E45; color: #ffffff; font-size: 12pt; border-radius: 10px; padding: 8px; }
            QPushButton:hover { background-color: #ff5555; }
        """)

        # Définie l'icone, le titre mais aussi le texte du contenu de la Box
        self.setIcon(QMessageBox.Question)
        self.setWindowIcon(QIcon(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/CN_Logo_Modified.png"))
        self.setWindowTitle("Reconfiguration")
        self.setText("\nÊtes-vous sûr de vouloir reconfigurer la clef, le token institutionnel et le chemin d'accès au dossier où sont enregistrés les rapports par défaut?")

        # Définie les boutons pour répondre à la question (role et mise en forme)
        self.buttonYes = self.addButton("Oui", QMessageBox.YesRole)
        self.buttonYes.setFont(QFont("Arial", 11, QFont.Bold))
        self.buttonNo = self.addButton("Non", QMessageBox.NoRole)
        self.buttonNo.setFont(QFont("Arial", 11, QFont.Bold))

        pixmap = QPixmap(os.path.dirname(os.path.abspath(__file__)) + "/../Logos/Config.svg")  # Chemin vers icône personnalisée
        self.setIconPixmap(pixmap.scaled(48, 48))  # Redimensionner l'icône et l'assigner
    

class InfoAPI(QDialog):
    """Classe de la boîte de dialogue sur les informations des API
    """
    def __init__(self, infos_API: dict):
        """Affiche les informations d'identification API dans une boite a onglets."""
        super().__init__()

        # Définie le titre et les dimensions de la Box
        self.setWindowTitle("Informations sur les API")
        self.resize(800, 250)

        # La structure principale est conçue d'un arrangement vertical qui contient un Widget avec 2 onglets (SciVal et Scopus)
        layout = QVBoxLayout()
        tab_main = QTabWidget()

        # L'onglet SciVal contient seulement l'onglet AuthorLookup
        tab_subScival = QTabWidget()
        tab_subScival.addTab(self._create_tab(infos_API['AuthorLookup']), "AuthorLookup")
        tab_main.addTab(tab_subScival, "SciVal")

        # L'onglet Scopus contient 3 onglets : AuthorSearch, AuthorRetrieval et CitationOverview
        tab_subScopus = QTabWidget()
        noms_API_Scopus = ["AuthorSearch", "AuthorRetrieval", "CitationOverview"]
        for i in range(len(noms_API_Scopus)):
            nom_API = noms_API_Scopus[i]
            tab_subScopus.addTab(self._create_tab(infos_API[nom_API]), nom_API)
        tab_main.addTab(tab_subScopus, "Scopus")

        # Ajout du Widget avec les 2 onglets pricipaux au Layout placé comme principale
        layout.addWidget(tab_main)
        self.setLayout(layout)

    def _create_tab(self, API: str):
        """Méthode utilitaire pour créer des onglets
        """
        tab = QWidget()
        tab_layout = QVBoxLayout()
        tab_layout.addWidget(self._create_info_widget(API)) # Le layout contient un QPlainTextEdit avec le contenu approprié
        tab.setLayout(tab_layout)

        return tab

    def _create_info_widget(self, infos):
        """Méthode utilitaire pour générer une zone de texte en fonction des infos de l'API en question fournies
        """
        info_widget = QPlainTextEdit()
        info_widget.setReadOnly(True)  # Pour empêcher l'édition du texte

        # Définie le fuseau horaire de Montréal
        montreal_timezone = pytz.timezone("America/Montreal")
        display_format = '%A %d %B %Y %H:%M:%S %Z'

        # Opérations ternaires qui attribue la conversion au fuseau horaire de Montréal si on a les informations des API
        formatted_date_GMT = "\t(" + str(pytz.timezone('GMT').localize(datetime.strptime(infos['Date'], "%a, %d %b %Y %H:%M:%S GMT")).astimezone(montreal_timezone).strftime(display_format)) + "/ à Montréal)" if infos['Date'] != 'None' else ''
        formatted_date_epoch = "\t\t\t(" + str(datetime.fromtimestamp(int(infos['X-RateLimit-Reset']), tz=pytz.UTC).astimezone(montreal_timezone).strftime(display_format)) + "/ à Montréal)" if infos['X-RateLimit-Limit'] != 'None' else ''

        # Affichage des informations
        info_widget.setPlainText("Date:\t\t" + infos['Date'] + formatted_date_GMT + "\nX-RateLimit-Limit:\t" + infos['X-RateLimit-Limit'] + "\nX-RateLimit-Remaining:\t" + infos['X-RateLimit-Remaining'] + "\nX-RateLimit-Reset:\t" + infos['X-RateLimit-Reset'] + formatted_date_epoch + "\n\nLes valeurs sont indiquées que lorsque vous avez utilisé l'API concernée, depuis l'ouverture du logiciel.")
        return info_widget


class Info(QDialog):
    """Classe de la boîte de dialogue sur les informations du logiciel
    """
    def __init__(self):
        """Affiche les informations sur le logiciel et ses credits."""
        super().__init__()

        # Définie le nom, redimensionne la Box et instancie le layout principal
        self.setWindowTitle("Informations sur AutoBib")
        self.resize(600, 600)
        layout = QVBoxLayout()

        info_widget = QTextEdit()
        info_widget.setReadOnly(True)  # Pour empêcher l'édition du texte
        
        info_widget.append("<b>AutoBib © Benjamin Lepourtois (2023)</b>")
        info_widget.append("")
        info_widget.append("<b>Pour citer ou reprendre ce logiciel :</b>")
        info_widget.append("Lepourtois. Benjamin. (2023). AutoBib [logiciel]. Montréal, QC : École de Technologie Supérieure.  Sous licence MIT.\n\n")

        info_widget.append("<b>Reprise de la bibliothèque publique de code :</b>")
        info_widget.append("<b>pybliometrics (version 3.5.2) :</b> https://github.com/pybliometrics-dev/pybliometrics")
        info_widget.append('<b>Article sur pybliometrics :</b> Rose, Michael E. and John R. Kitchin (2019): "pybliometrics: Scriptable bibliometrics using a Python interface to Scopus", SoftwareX 10 (2019) 100263')
        info_widget.append('\n')
        
        # Lire le contenu du fichier LICENSE et l'afficher dans la box à la suite
        with open(os.path.join(os.path.dirname(__file__), '..', 'LICENSE'), 'r') as fichier_licence:
            contenu_licence = fichier_licence.read()
        info_widget.append("<b>MIT license (at the root of the project) :</b>")
        info_widget.append(str(contenu_licence))
        
        # Ajout de la zone de texte créée au layout principal et identifie le layout principal comme réel layout principal de la Box
        layout.addWidget(info_widget)
        self.setLayout(layout)


class Timer:
    """Classe du chronomètre/timer utilisé pour mesurer le temps de la création d'une fiche bibliométrique
    """
    def __init__(self):
        """Initialise les compteurs internes du chronometre."""
        # Variables propres à l'objet
        self._start_time = None
        self._elapsed_time = 0
        self._timer_thread = None
        self._running = False
        self._stop_event = threading.Event()

    def _timer_function(self):
        """Méthode "privée" permettant de compter seconde après seconde
        """
        while not self._stop_event.is_set():
            time.sleep(1)
            self._elapsed_time += 1

    def start(self):
        """Méthode publique permettant de démarrer le chronomètre
        """
        if not self._running:
            self._start_time = time.time()
            self._running = True
            self._stop_event.clear()
            self._timer_thread = threading.Thread(target=self._timer_function)
            self._timer_thread.start()

    def stop(self):
        """Méthode publique permettant d'arrêter le chronomètre
        """
        if self._running:
            self._stop_event.set()
            self._timer_thread.join()
            self._running = False

    def reset(self):
        """Méthode publique permettant de remettre à zéro le chronomètre
        """
        self._start_time = None
        self._elapsed_time = 0

    def get_elapsed_time(self):
        """Méthode publique "getter" permettant d'obtenir le temps actuel du chronomètre
        """
        if self._start_time is None:
            return 0
        return self._elapsed_time + int(time.time() - self._start_time)
    
