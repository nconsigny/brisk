#########
# Read, Write and Modify xml (Brisk) files
# others files utils

import xml.etree.ElementTree as Etr
import subprocess
import os
import time
import glob
import Environnement as EnvB
import shutil
import re


def remplacer_nombres_dans_xml(fichier_entree, fichier_sortie, balise_debut, nouveaux_nombres):
    balise_fin = "END_".join(balise_debut)
    try:
        # Analyser le fichier XML d'entrée
        arbre = Etr.parse(fichier_entree)
        racine = arbre.getroot()

        # Trouver les éléments entre les balises de début et de fin
        # anciens_nombres = [] # Liste pour stocker les anciens nombres
        for elem in racine.iter():
            if elem.tag == balise_debut:
                # Stocker les anciens nombres dans la liste
                # anciens_nombres = [float(num) for num in elem.text.split(",")]
                # Remplacer le texte existant par les nouveaux_nombres
                elem.text = ",".join(str(num) for num in nouveaux_nombres)
            elif elem.tag == balise_fin:
                break

        # Enregistrer le XML modifié
        with open(fichier_sortie, "wb") as f:
            arbre.write(f)
        # print(f"XML modifié enregistré sous {fichier_sortie}")

        # Afficher les anciens nombres
        # print(f"Anciens nombres: {anciens_nombres}")
    except Exception as e:
        print(f"Erreur: {e}")


def remplacer_double_nombres_dans_xml(fichier_entree, fichier_sortie, balise_debut, n1, n2):
    # Créer une nouvelle liste pour stocker les valeurs intercalées
    n12 = []
    # Boucle pour intercaler les éléments de n1 et n2
    for i in range(len(n1)):
        n12.append(n1[i])
        n12.append(n2[i])
    balise_fin = "END_".join(balise_debut)
    try:
        # Analyser le fichier XML d'entrée
        arbre = Etr.parse(fichier_entree)
        racine = arbre.getroot()
        # Trouver les éléments entre les balises de début et de fin
        for elem in racine.iter():
            if elem.tag == balise_debut:
                elem.text = ",".join(str(num) for num in n12)
            elif elem.tag == balise_fin:
                break

        # Enregistrer le XML modifié
        with open(fichier_sortie, "wb") as f:
            arbre.write(f)
        # print(f"XML modifié enregistré sous {fichier_sortie}")
    except Exception as e:
        print(f"Erreur: {e}")


def remplacer_str_dans_xml(fichier_entree, fichier_sortie, balise_debut, new_str):
    balise_fin = "END_".join(balise_debut)
    try:
        # Analyser le fichier XML d'entrée
        arbre = Etr.parse(fichier_entree)
        racine = arbre.getroot()

        # Trouver les éléments entre les balises de début et de fin
        for elem in racine.iter():
            if elem.tag == balise_debut:
                elem.text = new_str
            elif elem.tag == balise_fin:
                break
        # Enregistrer le XML modifié dans un nouveau fichier
        with open(fichier_sortie, "wb") as f:
            arbre.write(f)
        # print(f"XML modifié enregistré sous {fichier_sortie}")
    except Exception as e:
        print(f"Erreur: {e}")


def lire_str_dans_xml(fichier_entree, balise_debut):
    balise_fin = "END_".join(balise_debut)
    try:
        # Analyser le fichier XML d'entrée
        arbre = Etr.parse(fichier_entree)
        racine = arbre.getroot()
        # Trouver les éléments entre les balises de début et de fin
        for elem in racine.iter():
            if elem.tag == balise_debut:
                return elem.text
            elif elem.tag == balise_fin:
                return ''
    except Exception as e:
        print(f"Erreur: {e}")


# Fonction pour remplacer une chaîne dans l'attribut 'description' des éléments 'fire'
def modif_fire_description_in_xml(file_path, new_string):
    try:
        # Analyser le fichier XML
        tree = Etr.parse(file_path)
        root = tree.getroot()

        # Trouver tous les éléments 'fire' et remplacer l'attribut 'description'
        for fire in root.iter('fire'):
            description = fire.get('description')
            if description is not None:
                fire.set('description', new_string)
                # Enregistrer le XML modifié dans le fichier
                with open(file_path, "wb") as f:
                    tree.write(f)
            else:
                print(f"erreur fire description non trouvée dans {file_path}")
    except Exception as e:
        print(f"Erreur: {e}")


def lire_nombres_xml(fichier_entree, balise_debut):
    anciens_nombres = []  # Liste pour stocker les anciens nombres
    balise_fin = "END_".join(balise_debut)
    try:
        # Analyser le fichier XML d'entrée
        arbre = Etr.parse(fichier_entree)
        racine = arbre.getroot()

        # Trouver les éléments entre les balises de début et de fin
        for elem in racine.iter():
            if elem.tag == balise_debut:
                # Stocker les anciens nombres dans la liste
                anciens_nombres = [float(num) for num in elem.text.split(",")]
            elif elem.tag == balise_fin:
                break
        print(f"nombres lus: {anciens_nombres}")
    except Exception as e:
        print(f"Erreur: {e}")
    return anciens_nombres


# Fonction récursive pour rechercher et remplacer la valeur
def find_and_replace_value_varname(element, varname, new_v):
    for child in element:
        if child.tag == 'varname' and child.text == varname:
            value_element = element.find('value')
            old_text = value_element.text
            value_element.text = str(new_v)
            return float(old_text)
        result = find_and_replace_value_varname(child, varname, new_v)
        if result is not None:
            return result
    return None


def find_and_replace_item_varname(element, varname, new_v, item):
    for child in element:
        if child.tag == 'varname' and child.text == varname:
            value_element = element.find(item)
            old_text = value_element.text
            value_element.text = str(new_v)
            return float(old_text)
        result = find_and_replace_item_varname(child, varname, new_v, item)
        if result is not None:
            return result
    return None


def find_and_replace_value_in_xml(file_path, varname_to_find, new_value):
    # Charger le fichier XML
    tree = Etr.parse(file_path)
    root = tree.getroot()

    # Fonction récursive pour rechercher et remplacer la valeur
    def find_and_replace_value(element, varname, new_v):
        for child in element:
            if child.tag == 'varname' and child.text == varname:
                value_element = element.find('value')
                old_text = value_element.text
                if new_value is not None:
                    value_element.text = new_v
                return float(old_text)
            result = find_and_replace_value(child, varname, new_v)
            if result is not None:
                return result
        return None

    # Appeler la fonction récursive sur la racine
    old_value = find_and_replace_value(root, varname_to_find, new_value)

    if old_value is not None:
        # Sauvegarder les modifications dans le fichier XML
        if new_value is not None:
            with open(file_path, "wb") as f:
                tree.write(f)
        return old_value
    else:
        print('Variable non trouvée')
        return float('nan')


# New EC5
# return value (string) or '?' if not find
def find_attribute_value(file_path, attribute):
    # Charger le fichier XML
    tree = Etr.parse(file_path)
    root = tree.getroot()

    for elem in root.iter():
        if attribute in elem.attrib:
            return elem.attrib[attribute]

    return '?'


def find_value_in_xml(file_path, varname_to_find):
    # Charger le fichier XML
    tree = Etr.parse(file_path)
    root = tree.getroot()

    # Fonction récursive pour rechercher la valeur
    def find_value(element, varname):
        for child in element:
            if child.tag == 'varname' and child.text == varname:
                value_element = element.find('value')
                old_text = value_element.text
                return float(old_text)
            result = find_value(child, varname)
            if result is not None:
                return result
        return None

    # Appeler la fonction récursive sur la racine
    old_value = find_value(root, varname_to_find)
    if old_value is not None:
        return old_value
    else:
        print('Variable non trouvée')
        return float('nan')


def lire_first_texte_arbre_xml(entree, balise_debut):
    balise_fin = "END_".join(balise_debut)
    str_lu = ''
    try:
        # Trouver les éléments entre les balises de début et de fin
        for elem in entree.iter():
            if elem.tag == balise_debut:
                str_lu = elem.text
                break
            elif elem.tag == balise_fin:
                break
        # print('texte lu:' + str_lu)
    except Exception as e:
        print(f"Erreur: {e}")
    return str_lu


def lire_texte_xml(fichier_entree, balise_debut):
    arbre = Etr.parse(fichier_entree)
    str_lu = lire_first_texte_arbre_xml(arbre.getroot(), balise_debut)
    return str_lu


def remove_files(chemin, f_extension):
    # Use glob.glob to find files with extension
    f_remove = glob.glob(os.path.join(chemin, f_extension))
    # supress each file with extension
    for f_to_remove in f_remove:
        os.remove(f_to_remove)
        print(f'Le fichier {f_to_remove} a été supprimé.')


def test_open_write(chemin_fichier):
    while True:
        """Vérifie si le fichier peut être ouvert en écriture."""
        if not os.path.exists(chemin_fichier):
            print(f"Le fichier '{chemin_fichier}' n'existe pas.")
            input("Appuyez sur une touche pour réessayer...")
            # Le fichier n'existe pas
        try:
            # Essaye d'ouvrir le fichier en mode écriture
            with open(chemin_fichier, 'a'):
                return True  # Le fichier peut être ouvert en écriture
        except Exception as e:
            print(f"Erreur fichier '{chemin_fichier}': {e}")
            input("Appuyez sur une touche pour réessayer...")


def brisk_test_subprocess():
    # Brisk.boucle_calculs()
    brisk_path = r'C:\Users\Francois.CONSIGNY\Documents\B-RISK 2023.1\BRISK.exe'
    model_paths = r'C:\Users\Francois.CONSIGNY\Documents\models\RISE_2'
    list_of_models = os.listdir(model_paths)

    for x in range(0, len(list_of_models)):
        model_folder_path = model_paths + '\\' + list_of_models[x]
        current_model = list_of_models[x]
        sub_model_folder_path = model_folder_path + '\\' + current_model
        # Erase all existing results
        remove_files(sub_model_folder_path, '*.csv')
        remove_files(sub_model_folder_path, '*.txt')
        remove_files(sub_model_folder_path, '*.pdf')
        remove_files(sub_model_folder_path, '*.xlsx')
        remove_files(sub_model_folder_path, 'output1.xml')
        # noinspection SpellCheckingInspection
        remove_files(sub_model_folder_path, 'dumpdata.dat')

        subprocess.Popen([brisk_path, model_folder_path])

        base_model_time = 0

        while not os.path.exists(sub_model_folder_path + '\\' + 'output1.xml'):
            # while the output isn't produced ye#t
            base_model_time = + 1  # just something to make the code wait for B-RISK
        while not os.path.exists(sub_model_folder_path + '\\' + current_model + '_results.xlsx'):
            # while the results isn't produced ye#t
            base_model_time = + 1  # just something to make the code wait for B-RISK
        while not os.path.exists(sub_model_folder_path + '\\' + current_model + '_zone.csv'):
            # while the zone isn't produced ye#t
            base_model_time = + 1  # just something to make the code wait for B-RISK
        print(base_model_time)

        time.sleep(5)  # just waiting to make sure no data is lost or corrupted

        os.system(r"taskkill /F /IM BRISK.exe")


def init_brisk_basemodel(model_gen):
    EnvB.ModelPaths = EnvB.CurPath + '\\' + EnvB.ModelName
    # Création du répertoire et Copie des fichiers .xlm et .dat
    if not os.path.exists(EnvB.ModelPaths):
        os.makedirs(EnvB.ModelPaths)  # Créer le répertoire de destination s'il n'existe pas
    else:
        shutil.rmtree(EnvB.ModelPaths)  # vide le répertoire de destination s'il existe
        os.makedirs(EnvB.ModelPaths)
    dest_dir = EnvB.ModelPaths + '\\' + f"{EnvB.BaseModel}0"
    os.makedirs(dest_dir)
    dest_dir = dest_dir + '\\' + f"{EnvB.BaseModel}0"
    os.makedirs(dest_dir)
    # model_gen = EnvB.CurPath + source_dir
    for file_name in os.listdir(model_gen):  # Parcourir les fichiers dans le répertoire source
        if file_name.endswith('.xml') or file_name.endswith('.dat'):
            # Construire les chemins complets source et destination
            source_file = os.path.join(model_gen, file_name)
            destination_file = os.path.join(dest_dir, file_name)
            # Copier le fichier
            shutil.copy2(source_file, destination_file)
            # print(f'Copié: {file_name}')
    return dest_dir


# TODO for precalculated CFAST Files see init_brisk_basemodel(model_gen)
def init_cfast_basemodel():
    EnvB.ModelPaths = EnvB.CurPath + '\\' + EnvB.ModelName
    # Création du répertoire
    if not os.path.exists(EnvB.ModelPaths):
        os.makedirs(EnvB.ModelPaths)  # Créer le répertoire de destination s'il n'existe pas
    else:
        shutil.rmtree(EnvB.ModelPaths)  # vide le répertoire de destination s'il existe
        os.makedirs(EnvB.ModelPaths)
    dest_dir = EnvB.ModelPaths + '\\' + f"{EnvB.BaseModel}0"
    os.makedirs(dest_dir)
    return dest_dir


# End Read, Write and Modify xml (Brisk) files
##########

#########
# Read, Write and Modify .in (Cfast) files
#

def update_cfast_field(filename, field, new_value):
    with open(filename, 'r') as file:
        content = file.read()

    # Construire la regex pour trouver le champ et le remplacer par la nouvelle valeur
    content = re.sub(f"{field} = [^,/\n]+", f"{field} = {new_value}", content)

    with open(filename, 'w') as file:
        file.write(content)

# End Read, Write and Modify .in (Cfast) files
##########
