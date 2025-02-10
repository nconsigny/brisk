import csv
import os
import sys
import pandas as pd
import socket
import shutil

import Files_utils


#########
# Read csv file

def extract_column_by_header(input_csv_file, header_text):
    with open(input_csv_file, mode='r') as file:
        csv_reader = csv.reader(file)
        header_row = next(csv_reader)  # Skip the header row
        column_index = header_row.index(header_text)
        column_data = []

        for row in csv_reader:
            column_data.append(row[column_index])

    return column_data


def read_csv_to_list(filename):
    """
    Lit un fichier CSV et crée une liste de chaînes de caractères.
    Args :
        filename (str) : Nom du fichier CSV.
    Returns :
        list : Liste de chaînes de caractères.
    """
    while True:
        """Vérifie si le fichier peut être ouvert en écriture."""
        if not os.path.exists(filename):
            print(f"Le fichier '{filename}' n'existe pas.")
            input("Appuyez sur une touche pour réessayer...")
            # Le fichier n'existe pas
        try:
            df = pd.read_csv(filename)
            # titres_colonnes = df['Titres de colonnes'].tolist()
            brisk_names = df.iloc[:, 0].tolist()  # Noms de la première colonne du CSV
            cfast_names = df.iloc[:, 1].tolist()  # Noms de la première colonne du CSV
            mapping_dict = dict(zip(cfast_names, brisk_names))
            # Renommer les colonnes du DataFrame current_data.rename (columns=mapping_dict, inplace=True)
            # Ferme le fichier
            with open(filename, 'r') as f:
                f.close()
            return brisk_names, mapping_dict
            # return titres_colonnes
        except Exception as e:
            print(f"Erreur fichier '{filename}': {e}")
            input("Appuyez sur une touche pour réessayer...")


# End Read csv file
#########

def unit_round(cherche: str) -> int:
    return next((item[1] for item in list_arrondis_unites if item[0] == cherche), None)


def u_round(x: float, u: str) -> float:
    n = unit_round(u)
    if n is not None:
        return round(x, unit_round(u))
    else:
        return x


#########
# local path and environment management

# List of path & directory for files & calculation
# Récupérer le nom de l'ordinateur
# nom_ordinateur = os.uname().#nodename
# Récupérer le login de l'utilisateur
# login_utilisateur = os.#getlogin()
# Récupérer le répertoire utilisateur
# repertoire_utilisateur = os.path.expanduser("~")
# CurPath = os.#getcwd()

Machine = socket.gethostname()
ProgPath = os.path.dirname(os.path.abspath(sys.argv[0]))
ProgFile = os.path.basename(sys.argv[0])
ProgName = ProgPath + '\\' + ProgFile
CalculationTime = 0
# noinspection SpellCheckingInspection
CfastPath = r'C:\Program Files\firemodels\cfast7\cfast'
if Machine == 'Fanfan2':  # if personal station FC
    BriskPath = r'C:\Program Files (x86)\BRANZ\B-RISK 2024.31\BRISK.exe'
    CstbGroup = r'C:\Users\franc\CSTBGroup\These_Francois_Consigny - Documents'
    UserPath = r'C:\Users\franc'
else:
    if Machine == 'Fanfan1':
        BriskPath = r'C:\Program Files (x86)\BRANZ\B-RISK 2024.3\BRISK.exe'
        CstbGroup = r'C:\Users\franc\CSTBGroup\These_Francois_Consigny - Documents'
        # CurPath = r'D:\Donnees\Thèse\Calculs Feu\Programmes\Models'
        UserPath = r'C:\Users\franc'
    else:
        if Machine == 'CSM-LOC010499':  # if station FC CSTB
            BriskPath = r'C:\Program Files (x86)\BRANZ\B-RISK 2024.1\BRISK.exe'
            CstbGroup = r'C:\Users\francois.consigny\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
            # CurPath = (CstbGroup + r'\DSSF_Projet_These_Francois_Consigny - Documents\Calculs Feu\Programmes\Models')
            UserPath = r'C:\Users\Francois.CONSIGNY'
        else:  # if station FC ENPC
            BriskPath = r'C:\Program Files (x86)\BRANZ\B-RISK 2024.3\BRISK.exe'
            CstbGroup = r'C:\Users\Francois.CONSIGNY\CSTBGroup\These_Francois_Consigny - Documents'
            # CurPath = r'D:\Donnees\Thèse\Calculs Feu\Programmes\Models'
            UserPath = r'C:\Users\Francois.CONSIGNY'

Brisk_already_write = ['McGregor 1', 'McGregor 3', 'USDA 4', 'USDA 5']
#########
# Choose contribution law
# ## Summary of methods: included in Contribution_Brisk_fin
# LG: Original method of LG and CD as in project Contrib_Brisk1 and Contrib_Brisk2
# (Brisk2 same as Brisk1 but with allow_contribution_protected-> True and allow_wind-> True)
# Flux : Replace contribution law with flux from Brisk instead of from gas Temperature as in project Contrib_Brisk3
# Mlr : Replace contribution law with MLR* as in project Contrib_Brisk4
# EC5 : According new version of EC5-1-2 Annex A
# ## End methods summary
contribution_methods = ['LG', 'Flux', 'Mlr', 'EC5']
# 'B' for Fire Safety Challenges of Tall Wood Buildings – Phase 2: Task 4 Engineering Methods (D.Brandon)
# 'E' for Eurocode (EC5-1-2 & EC1-1-2 Annex A)
parametric_curves = 'B'

default_method = contribution_methods[0]
# print(f"La méthode par défaut est maintenant : {default_method}")

zone_models = [('B', 'Brisk'), ('BT', 'Brisk'), ('C', 'Cfast'), ('E1', 'PC'), ('E2', 'PC')]
# TODO function choose zone_model & iteration if several ?
# manual choose of zone_model
current_zone_model = zone_models[0]
# global variables initialised by init_environnement_from_method
ModelName = ModelPaths = CurPath = Thermal_file = mlr_file = courbe_iso_file = Results_file = BaseModel = ''

if current_zone_model[1] == 'PC':
    max_iter = 40
else:
    max_iter = 20  # TODO >10 PROVISOIRE à modifier ?
min_humidity_brisk = 0.5  # TODO Check sensibility to this value (convergence??)
default_wood_density = 460  # EC5-1-2 A.3 Note 1
experiment_excel_file = ''  # Fichier de base d'extraction des essais
current_experiment_excel_file = ''  # Fichier courant de calcul
result_list_char = []  # Liste des noms d'essais du fichier experiment_excel_file pour écriture des résultats (char)

# ### End of List of path & directory for files & calculation

# # List of global variable
list_arrondis_unites = [('[mm]', 0), ('[mn]', 0), ('[mm.]', 1), ('[mm/min]', 3), ('[m²]', 4), ('%', 5), ('[m]', 3),
                        ('[kg]', 3), ('[MJ/m²]', 0), ('[°C]', 1), ('[MW]', 1), ('[m1/²]', 5), ('[m/s]', 3), ('[s]', 1),
                        ('[kW]', 1), ('[m1/2]', 3), ('[°]', 1), ('[]', 3), ('[mn.]', 1), ('[J/m²Ks1/2]', 0),
                        ('[kJ/g]', 3), ('[kW/m²]', 0)]
liste_lue = []
long_line = [ProgName]
mapping_cfast_brisk = dict()
# Default values for HRR fuel calculation regarding ref. [1] & [2] (see introduction in main.py
alpha = 0.012  # [kW/s²] Fire growth in alpha*t² Brandon: 0.047
alpha1 = 0.4  # [kg/(s m5/2)] Brandon: 0.4
alpha2 = 3.01E+06  # [Ws/kg] Brandon: 3.01E+06
alpha3 = 0  # not used
alpha4 = 0.1  # [-] Brandon: 0.1
alpha5 = 0.8  # [-] 80% of total mass loss  Brandon: 0.8
alpha6 = 0.6  # [-] Begin of decay phase at 60% of movable fuel consumed Brandon: 0.5
alphas = [alpha, alpha1, alpha2, alpha3, alpha4, alpha5, alpha6]
default_total_time = 7500  # [s] t_fin
if current_zone_model[1] == 'Cfast':
    default_excel_interval = 100
    default_hrr_interval = 100  # []
else:
    default_excel_interval = 60  # [] Warning /!\ could vary cf. email C.Wade 26/11/2024
    default_hrr_interval = 50  # []
default_leak_area_ratio = 0.005  # [m²/m²]
default_ceiling_nodes = 15  # []
default_wall_nodes = 15  # []
default_floor_nodes = 10  # []
# Structural wood heat of combustion [kJ/g] New EC5: EC5-1-2 A.3 Note 1: 17.5
# Brisk see dbase/thermal.mdb : 14 [kJ/g]
# TODO Store in Material and Compartment and make method structural_wood_heat_of_combustion ?
default_structural_wood_heat_of_combustion = 12.4  # Cribs Value for B_Risk
default_wood_ignition_temp = 250  # [°C] Not used New EC5 (to cap wood contribution law)
default_ceiling_factor = 0.75  # [] 0.75 regarding original LG article
# The analysis of the tests carried out by [5] (Test 2) and [7] (Tests 2, 3, and 5) shows that the average value of
# the char depth measured at the ceiling is between 0.6 and 0.75 of the average char depth measured at the walls.
allow_char_energy_storage = False  # New EC5: EC5-1-2 A.3 Note 2 alpha_st take as RR_UL & RR_LL from LG article if True,
# 1 if False
default_pressure = 101325  # [Pa]
# noinspection SpellCheckingInspection
default_smokeview_step = 15  # []
allow_contribution_protected = True  # Version Brisk1; Brisk2 -> True
allow_wind = True  # Version Brisk1; Brisk2 -> True
fall_off_temp = 400  # [°C]
flashover_temp = 500  # [°C] Default value of UL temp (500 default Brisk criterion)
ingberg_temp_seuil = 150  # [°C]


# ### End of List of global variable

def choose_method():
    global contribution_methods, parametric_curves, allow_contribution_protected, current_zone_model, \
        allow_char_energy_storage
    if current_zone_model[1] == 'PC':
        contribution_methods = ['EC5']
        allow_char_energy_storage = False  # TODO implement allowing ?
        allow_contribution_protected = False  # TODO implement allowing ?
        index = 0
        while index not in ['1', '2']:
            try:
                index = input("Tapez:\n "
                              "1 pour la méthode Eurocode (EC5-1-2 & EC1-1-2 Annex A) \n "
                              "2 pour la méthode Fire Safety Challenges of Tall Wood Buildings – Phase 2: "
                              "Task 4 Engineering Methods (D.Brandon): ")
                if index == '1':
                    parametric_curves = 'E'
                    current_zone_model = zone_models[2]
                elif index == '2':
                    parametric_curves = 'B'
                    current_zone_model = zone_models[3]
                else:
                    print(f"Erreur : tapez 1 ou 2.")
            except ValueError:
                print("Erreur : Veuillez entrer un nombre entier valide.")
    else:
        methods = contribution_methods.copy()
        contribution_methods = []
        print('Choix de la méthode de contribution, tapez:')
        for i in methods:
            print(f"{methods.index(i)}- {i}")
        print('all- Toutes')
        print(f"Return- Exit")
        index = 0
        # Boucle while pour demander une entrée valide
        while 0 <= index < len(methods):
            try:
                index = input("Veuillez taper le nombre de la methode à ajouter ou return "
                              "si la liste de methode est complète  : ")
                if index == '':
                    return
                elif index == 'all':
                    contribution_methods = methods
                    return
                else:
                    index = int(index)
                if 0 <= index < len(methods):
                    contribution_methods.append(methods[index])
                    print(f"Methode de contribution:", contribution_methods)
                else:
                    print(f"Erreur : Le nombre doit être entre 0 et {len(methods) - 1}.")
            except ValueError:
                print("Erreur : Veuillez entrer un nombre entier valide.")


# end of local path and environment management
#########

def init_environnement_from_method(method):
    global contribution_methods, default_method, default_excel_interval, default_hrr_interval
    global Results_file, CurPath, BaseModel, ModelName, ModelPaths, long_line
    global Thermal_file, mlr_file, courbe_iso_file, experiment_excel_file, liste_lue, mapping_cfast_brisk

    default_method = method
    # print(f"La méthode par défaut est maintenant : {default_method}")

    # DO check & copy precalculated Brisk files in CurPath & change .xlsx filenames with prefix B_date ?
    Results_file = f" Recap Results_{current_zone_model[1]}" + default_method + '.xlsx'
    CurPath = UserPath + f"\\{current_zone_model[0]}_Models_" + default_method
    BaseModel = 'basemodel_' + default_method + '_iter'
    if not os.path.exists(CurPath):
        os.makedirs(CurPath)  # Créer le répertoire de destination s'il n'existe pas
    dest_dir = CurPath + r'\input'
    if not os.path.exists(dest_dir):
        source_dir = CstbGroup + r'\Calculs Feu\Programmes\python\input'
        shutil.copytree(source_dir, dest_dir, dirs_exist_ok=True)   # Warning, override existing files in dest_dir
    # if current_zone_model[1] == 'Brisk':
    if current_zone_model[0] == 'B':
        for name_experiment in Brisk_already_write:
            ModelName = name_experiment
            ModelPaths = CurPath + '\\' + ModelName
            dest_dir = Files_utils.init_brisk_basemodel(
                CstbGroup + r'\Calculs Feu\Programmes\python\Brisk_pre_calculated' + '\\' + name_experiment)
            os.rename(dest_dir + '\\' + 'basemodel_Iter0.xml', dest_dir + '\\' + f"{BaseModel}0.xml")
    ModelName = 'Temp'
    ModelPaths = CurPath + '\\' + ModelName
    Thermal_file = CurPath + r'\input\Thermal.xlsx'
    # noinspection SpellCheckingInspection
    mlr_file = CurPath + r'\input\241113 Function MLR_echar.xlsx'
    courbe_iso_file = CurPath + r'\input\241209 Courbe_ISO.xlsx'
    # liste_lue = read_csv_to_list(CurPath + r'\input\Brisk_result_keep_header.csv')
    liste_lue, mapping_cfast_brisk = read_csv_to_list(CurPath + r'\input\result_keep_header.csv')
    # experiment_excel_file = CurPath + r'\input\241108 Recap Experiment_these.xlsx'  # Fichier de Base
    if current_zone_model[0] == 'BT':
        experiment_excel_file = CurPath + r'\input\250205 Dwellings_distribution.xlsx'  # Fichier de Base
    else:
        experiment_excel_file = CurPath + r'\input\241108 Recap Experiment_these.xlsx'  # Fichier de Base
    if current_zone_model[1] == 'Cfast':
        default_excel_interval = 100
        default_hrr_interval = 100  # []
    else:
        default_excel_interval = 60  # [] Warning /!\ could vary cf. email C.Wade 26/11/2024
        default_hrr_interval = 50  # []
