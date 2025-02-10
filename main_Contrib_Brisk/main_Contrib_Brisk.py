#######################################################################################################################
# Introduction
#######################################################################################################################

#########
# Author : François CONSIGNY | https://www.linkedin.com/in/fran%C3%A7ois-consigny-a7b4115b/
#########

#########
# Used integrated development environment (IDE) for programming in Python
#         PyCharm 2024.1.6 (Community Edition)
#         Build #PC-241.19072.16, built on August 8, 2024,
#         Runtime version: 17.0.11+1-b1207.30 amd64
#         VM: OpenJDK 64-Bit Server VM by JetBrains s.r.o.
#         Windows 11.0
#         GC: G1 Young Generation, G1 Old Generation
#         Memory: 4096M
#         Cores: 32
#         Registry:
#           ide.experimental.ui=true
#########

#########
# Overall program scope, objective & method.
#   The objective of this project is to develop an accessible and
#   rapid calculation tool that will enable the conduct of parametric studies on fire compartment with exposed mass
#   timber.
#   The project is based on recent works by the Université Laval [1] and RISE [2], and proposes improvements
#   to the method and the use of verified open access two-zone models
#   (B-Risk developed by BRANZ or CFast developed by NIST).
#   The global process is managed by a Python program, which oversees both pre- and post-processing operations
#   as follows:
#           1.	The process is initiated by calculating the heat release rate (HRR) of the fuel load, regarding the
#           calorific load density of the movable fuel in the compartment.
#           2.	The results of the time-history curves of temperature, hot layer height, oxygen concentration
#           and incident flux on boundary elements are calculated, taking into account the HRR and
#           the geometrical properties of the compartment with the zone model (BRISK or CFAST).
#           3.	The charring rate of exposed CLT (Cross Laminated Timber) is then calculated for each time step ,
#           based on the aforementioned values.
#           The contribution of timber to the heat load is then deducted from this charring rate.
#           4.	The combustion energy released from the timber is then added to the heat release rate (HRR) of step 1,
#           resulting in a new, increased HRR.
#           5.	Subsequently, an iterative process is initiated, repeating steps 1 to 4 with the new
#           increased HRR, and continuing until the convergence criteria is reached (i.e. the difference between HRR of
#           iteration n and iteration n-1 becomes negligible).
#           6.	Ultimately, the static load capacity of the CLT in accordance with the residual section of CLT
#           (or alternatively, through a thermo-mechanical calculation with the final time history temperature applied
#           to the CLT) can be evaluated.
#
# References:
# [1] Fire Dynamics of Mass Timber Compartments with Exposed Surfaces: Development of an Analytical Model |
#       L.Girompaire & C.Dagenais |
#       https://link.springer.com/article/10.1007/s10694-023-01528-y

# [2] Predictive method for fires in CLT and glulam structures – A priori modelling versus real scale compartment
#       fire tests & an improved method |
#       D.Brandon |
#       https://www.diva-portal.org/smash/get/diva2:1075003/FULLTEXT01

# Contrib_Brisk_fin : Final version with implementation of different methods of contribution
# See Environnement.Methods
# for MLR* method see
# [3] An Empirical Correlation for Burning of Spruce Wood in Cone Calorimeter for Different Heat Fluxes |
#       P.Lardet & A.Coimbra |
#       https://link-springer-com.extranet.enpc.fr/content/pdf/10.1007/s10694-024-01603-y.pdf
#
#######################################################################################################################
# End Introduction
#######################################################################################################################

import Brisk_calc as Contrib_calc
import Environnement as EnvB
import pandas as pd
import time
from datetime import timedelta
import sys
from Files_utils import init_cfast_basemodel
import subprocess
# # from concurrent.futures import ProcessPoolExecutor
# from concurrent.futures import ThreadPoolExecutor
# import multiprocessing


class DoubleOutput:
    def __init__(self, fichier):
        self.console = sys.stdout
        self.fichier = fichier

    def write(self, message):
        self.console.write(message)
        self.fichier.write(message)

    def flush(self):
        self.console.flush()
        self.fichier.flush()


def fill_none_with_previous(df, header):
    """
    Remplit les valeurs None dans une colonne spécifiée avec la valeur de la première ligne précédente non None.
    :param df : DataFrame pandas
    :param header : Nom de la colonne à vérifier et remplir
    :return : DataFrame modifié
    """
    if header in df.columns:
        # Parcourir chaque ligne de la colonne spécifiée
        for i in range(1, len(df)):
            if pd.isna(df.iloc[i][header]):
                # Trouver la première valeur non None précédente
                for j in range(i - 1, -1, -1):
                    if not pd.isna(df.iloc[j][header]):
                        df.iloc[i, df.columns.get_loc(header)] = df.iloc[j][header]
                        break
    else:
        print(f"La colonne '{header}' n'existe pas dans le DataFrame.")

    return df


# def process_element(elem):
#     """Traite un seul élément de la liste"""
#     if EnvB.current_zone_model[1] == 'Brisk':
#         if elem.id not in EnvB.Brisk_already_write:
#             elem.write_brisk_base_model()
#         else:
#             elem.write_brisk_default_values()
#     elif EnvB.current_zone_model[1] == 'Cfast':
#         EnvB.ModelName = f"{elem.id}"
#         dest_dir = init_cfast_basemodel()
#         elem.to_cfast(dest_dir, 0)
#     Contrib_calc.boucle_calculs(elem)


# Launch the Windows Brisk program from WSL
def launch_windows_brisk():
    windows_brisk_path = "C:\\Program Files\\Brisk\\Brisk.exe"  # update this path if needed
    try:
        subprocess.Popen(["cmd.exe", "/c", "start", "", windows_brisk_path])
        print("Launched Brisk program on Windows.")
    except Exception as e:
        print(f"Failed to launch Brisk program: {e}")


def main():
    launch_windows_brisk()
    # noinspection SpellCheckingInspection
    # #Déclarer le fichier comme une variable globale
    # fichier = open(EnvB.CurPath + '\\run.txt', 'w')
    # double_output = DoubleOutput(fichier)
    # sys.stdout = double_output

    # Choose experiment list
    essai = 'a'
    essais = []
    while not essai == '' and not essai == 'all':
        essai = input("liste d'essais ?")
        if not essai == '':
            essais.append(essai)
        print(essais)
    EnvB.choose_method()
    # Or, manually:
    # EnvB.contribution_methods = ['LG']

    for item in EnvB.contribution_methods:  # part of ['LG', 'Flux', 'Mlr', 'EC5'] (by choose_method)
        EnvB.init_environnement_from_method(item)
        Contrib_calc.init_courbe_iso_and_mlr()
        # Déclarer le fichier comme une variable globale
        fichier = open(EnvB.CurPath + f"\\run_{EnvB.default_method}.txt", 'w')
        double_output = DoubleOutput(fichier)
        sys.stdout = double_output
        start_time = time.time()
        print("Heure de début:", time.strftime("%H:%M:%S", time.localtime(start_time)))
        c_list = Contrib_calc.init_experiment_list(essais)[1]
        for elem in c_list:
            if EnvB.current_zone_model[1] == 'Brisk':
                if elem.id not in EnvB.Brisk_already_write:  # ['McGregor 1', 'McGregor 3', 'USDA 4', 'USDA 5']:
                    # List of experiment with Basemodel already generated with Brisk
                    # (McGregor Propane Fuel, USDA sprinkler)
                    # TODO maj heat_of_combustion (EC5) for USDA4,5 ?
                    elem.write_brisk_base_model()
                # New EC5
                else:
                    elem.write_brisk_default_values()
            elif EnvB.current_zone_model[1] == 'Cfast':
                EnvB.ModelName = f"{elem.id}"
                dest_dir = init_cfast_basemodel()
                elem.to_cfast(dest_dir, 0)
            Contrib_calc.boucle_calculs(elem)

        # # Utiliser le nombre de cœurs disponibles moins 1 pour laisser un cœur libre
        # max_workers = 2  # max(1, multiprocessing.cpu_count() - 1)
        #
        # # with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # # with ProcessPoolExecutor(max_workers=max_workers) as executor:
        # with ThreadPoolExecutor(max_workers=max_workers) as executor:
        #     # with multiprocessing.Pool(processes=max_workers) as pool:
        #     # Exécuter le traitement en parallèle
        #     list(executor.map(process_element, c_list))

        # TODO: balayer les projets à calculer à partir de
        #  fichiers de calculs existants (par Exemple NRC LS 5 qui ne peut être lu à partir du fichier excel) + les
        #  projets dont les fichiers de calculs sont générés à partir du fichier Excel ou pas ...
        #  (peut-être plus simple de traiter ce cas "à la mano ?")

        end_time = time.time()
        elapsed_time = end_time - start_time
        # Convertir en heures, minutes et secondes
        elapsed_time_formatted = str(timedelta(seconds=round(elapsed_time)))
        print(f"Temps de calcul: {elapsed_time_formatted}")
        if len(c_list) > 0:
            moyenne: float = round(elapsed_time / len(c_list))
            print(f"Temps par essai = {moyenne}")
        # Restaurer la sortie standard et fermer le fichier à la fin du programme
        sys.stdout = sys.__stdout__
        fichier.close()
    # sys.stdout = sys.__stdout__
    # fichier.close()


if __name__ == "__main__":
    main()
