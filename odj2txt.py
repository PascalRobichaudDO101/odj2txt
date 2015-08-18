#!c:\python27\python.exe
# coding: utf8

"""Extraction des contrats de fichiers d'ordre du jour

Convertit la section 20 - Affaires contractuelles du fichier PDF 
de l'ordre du jour en format texte.

Version 1.1, 2015-08-18
Développé en Python 3.4
Licence CC-BY-NC 4.0 Pascal Robichaud, pascal.robichaud.do101@gmail.com
"""

import pdf2txt
import os
import csv


def main():
    """Partie principale du traitement.

    Transforme tous les fichiers PDF dans le répertoire désigné 
    en fichiers texte.
    """
    print
    print("Debut du traitement")
    print

    # Répertoire où les fichiers PDF sont enregistrés
    REPERTOIRE_PDF = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\PDF"
    
    if not os.path.exists(REPERTOIRE_PDF):
        raise ValueError("Le repertoire " + REPERTOIRE_PDF + " pour les fichier PDF n'existe pas.")

    # Répertoire où le fichier texte résultant sera sauvegardé
    REPERTOIRE_TXT = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\TXT"
    
    if not os.path.exists(REPERTOIRE_TXT):
        os.makedirs(REPERTOIRE_TXT)
    
    for filename in os.listdir(REPERTOIRE_PDF):
        fichier_PDF = os.path.join(REPERTOIRE_PDF, filename)
        print("Traitement du fichier %s" % fichier_PDF)
        transforme_pdf_en_txt(fichier_PDF, REPERTOIRE_TXT)

    print
    print("Fin du traitement")


def transforme_pdf_en_txt(fichier_PDF, REPERTOIRE_TXT):
    
    # Titre de la section où se retrouvent les contrats.
    TITRE_SECTION_20 = " Affaires contractuelles"

    # Titre de la section suivant celle où se retrouvent les contrats.
    TITRE_SECTION_30 = " Administration et finances"

    prefixe_txt = os.path.splitext(os.path.basename(fichier_PDF))[0]
    fichier_TXT_temp = os.path.join(REPERTOIRE_TXT, prefixe_txt + '_temp.txt')
    fichier_TXT = os.path.join(REPERTOIRE_TXT, prefixe_txt + '.txt')
    odj_traites = open(fichier_TXT, "w")

    est_dans_section_affaires_contractuelles = False
    est_dans_section_suivante = False
    compteur_page = 0

    while not est_dans_section_suivante:
        compteur_page += 1

        print("Traitement de la page %s" % compteur_page)
        
        args = [
            'pdf2txt',
            '-p', str(compteur_page),
            '-o', fichier_TXT_temp,
            fichier_PDF,
        ]
        
        pdf2txt.main(args)

        with open(fichier_TXT_temp, 'r') as f:
            
            for ligne in f:

                if not est_dans_section_affaires_contractuelles:
                    if TITRE_SECTION_20 in ligne:
                        est_dans_section_affaires_contractuelles = True

                if TITRE_SECTION_30 in ligne:
                    est_dans_section_suivante = True
                    break
                    
                elif est_dans_section_affaires_contractuelles:

                    if ligne.startswith("['Page "):
                        # Ne pas écrire le numéro de page du pied-de-page
                        break
                    else:
                        # Ajouter la ligne dans le fichier fichier_TXT
                        odj_traites.writelines(ligne)
    
    os.remove(fichier_TXT_temp)
                        
    odj_traites.close()

if __name__ == '__main__':
    main()
