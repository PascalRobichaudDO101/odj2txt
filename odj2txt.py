#!c:\python27\python.exe
# coding: utf8

"""Extraction des contrats de fichiers d'ordre du jour

Convertit la section 20 - Affaires contractuelles du fichier PDF de l'ordre
du jour en format texte.

Version 1.0, 2015-08-18
Développé en Python 3.4
Licence CC-BY-NC 4.0 Pascal Robichaud, pascal.robichaud.do101@gmail.com
"""

import pdf2txt
import os
import csv


def main():
    """Partie principale du traitement.

    Transforme tous les fichiers PDF dans le répertoire désigné en fichiers
    texte.
    """
    print
    print("Debut du traitement")
    print

    # Répertoire où les fichiers PDF sont enregistrés.
    #
    # C'est vraiment plus simple de mettre les commentaires sur la ligne
    # au-dessus plutôt que s'essayer de garder aligné sur une colonne
    # imaginaire. Le commentaire est utile parce que c'est quelque chose
    # qu'un utilisateur du script pourrait vouloir changer.
    REPERTOIRE_PDF = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\PDF"

    # Répertoire où le fichier texte résultant sera sauvegardé.
    REPERTOIRE_TXT = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\TXT"

    # J'aime bien quand le code qu'on lit est une représentation de sa
    # description. Si on compare le commentaire de documentation au début de
    # la fonction et le code on devrait pouvoir voir le lien d'un à l'autre.
    for filename in os.listdir(REPERTOIRE_PDF):
        fichier_PDF = os.path.join(REPERTOIRE_PDF, filename)
        print("Traitement du fichier %s" % fichier_PDF)
        transforme_pdf_en_txt(fichier_PDF, REPERTOIRE_TXT)

    print
    print("Fin du traitement")


def transforme_pdf_en_txt(fichier_PDF, REPERTOIRE_TXT):
    # Titre de la section où se retrouvent les contrats.
    #
    # Ces informations ne sont utiles que pour la transformation d'un fichier,
    # alors ça a du sens de le mettre ici plutôt que dans main().
    TITRE_SECTION_20 = " Affaires contractuelles"

    # Titre de la section suivant celle où se retrouvent les contrats.
    TITRE_SECTION_30 = " Administration et finances"

    # OK, je suis pas trop fier de cette ligne-là...
    prefixe_txt = os.path.splitext(os.path.basename(fichier_PDF))[0]
    fichier_TXT = os.path.join(REPERTOIRE_TXT, prefixe_txt + '.txt')
    odj_traites = open(fichier_TXT, "w")
    fodj_traites = csv.writer(odj_traites, delimiter=';')

    # Les noms de variables sont assez explicites pour ne pas avoir besoin de
    # plus de commentaires.
    est_dans_section_affaires_contractuelles = False
    est_dans_section_suivante = False
    compteur_page = 0

    while not est_dans_section_suivante:
        compteur_page += 1

        print("Traitement de la page %s" % compteur_page)

        # Idéalement tu veux éviter de répéter une information
        # plusieurs fois dans le code pour prévenir des problèmes
        # de synchronisation lors de changements. Aussi, tu peux
        # utiliser plusieurs lignes pour améliorer la lisibilité.
        nom_du_fichier_page = os.path.join(
            REPERTOIRE_TXT,
            '%s_page_%d.txt' % (prefixe_txt, compteur_page))
        args = [
            'pdf2txt',
            '-p', str(compteur_page),
            '-o', nom_du_fichier_page,
            fichier_PDF,
        ]
        pdf2txt.main(args)

        with open(nom_du_fichier_page, 'r') as f:
            # Pourquoi du CSV? Tu peux lire les lignes directement de
            # l'objet retourné par open().
            reader = csv.reader(f, delimiter='|')

            for ligne in reader:
                # Je suis pas trop certain de comprendre pourquoi y'a un
                # détour d'encodage (ça devrait probablement être
                # documenté ou enlevé !)
                #
                # ligne = ligne.encode('utf-8')

                # Tu peux utiliser une expression plus simple et directe
                # pour évaluer True/False que `== False`.
                if not est_dans_section_affaires_contractuelles:
                    if TITRE_SECTION_20 in ligne:
                        est_dans_section_affaires_contractuelles = True

                if TITRE_SECTION_30 in ligne:
                    est_dans_section_suivante = True
                    break
                elif est_dans_section_affaires_contractuelles:
                    # Vaut mieux ne pas réinventer la roue et utiliser
                    # les fonctions de base, comme ça je n'ai pas
                    # besoin d'apprendre une nouvelle fonction à chaque
                    # fois que je lis un nouveau code :-)
                    if ligne.startswith("['Page "):
                        # Ne pas écrire le numéro de page du pied-de-page.
                        break
                    else:
                        # Ajouter la ligne dans le fichier fichier_TXT.
                        fodj_traites.writerow(ligne)

    # Pas besoin de fermer explicitement le fichier avec `with`.

    odj_traites.close()


if __name__ == '__main__':
    main()
