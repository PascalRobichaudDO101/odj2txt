#!c:\python27\python.exe
# coding: utf8

"""Convertir la section 20 - Affaires contractuelles du fichier PDF de l'ordre du jour en format texte
Version 2.4, 2015-08-18
Développé en Python 3.4
Licence CC-BY-NC 4.0 Pascal Robichaud, pascal.robichaud.do101@gmail.com"""

import pdf2txt
import os
import csv
     
#Fonction left
def left(s, amount = 1, substring = ""):

    if (substring == ""):
        return s[:amount]
    else:
        if (len(substring) > amount):
            substring = substring[:amount]
        return substring + s[:-amount]

#Début du traitement 
def main():        
    print
    print("Debut du traitement")
    print

    REPERTOIRE_PDF = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\PDF"              #Répertoire où les fichiers PDF sont enregistrés#
    fichier_PDF = ""                                                            #Nom du fichier PDF traité
    REPERTOIRE_TXT = "C:\\ContratsOuvertsMtl\\Ordres_du_jour\\TXT"              #Répertoire où le fichier texte résultant sera sauvegardé
    fichier_TXT = ""                                                            #Nom du fichier texte qui sera généré

    TITRE_SECTION_20 = " Affaires contractuelles"                               #Titre de la section où se retrouvent les contrats
    TITRE_SECTION_30 = " Administration et finances"                            #Titre de la section suivant celle où se retrouvent les contrats

    est_dans_section_affaires_contractuelle = False                             #Variable pour savoir si on est rendu à la section des contrats, pour ne pas sauvegarder
                                                                                #les premières pages inutilement
    continuer = True                                                            #Variable pour arrêter le traitement une fois que la section des contrats est terminée                                                                                
                                                                                
    compteur_page = 0                                                           #Compteur pour le traitement des pages       

    for filename in os.listdir(REPERTOIRE_PDF):                                 #Passer au travers des fichiers PDF

        fichier_PDF = REPERTOIRE_PDF + "\\" + filename
        fichier_TXT = REPERTOIRE_TXT + "\\" + filename.replace("pdf","txt")

        #Ouverture du fichier fichier_TXT pour sauvegarder le traitement
        odj_traites = open(fichier_TXT, "w")      
        fodj_traites = csv.writer(odj_traites, delimiter = ';') 
      
        while continuer:                                                        #Passer au travers des pages du fichier PDF
        
            compteur_page = compteur_page + 1                                   #Compteur pour le traitement des pages
            
            print("Traitement de la page %s" % compteur_page)                   #Afficher le numéro de la page comme indicateur que le traitement fonctionne
            
            arg = ["", '-p', '' + str(compteur_page) + '', '-o', 'C:\\ContratsOuvertsMtl\\Ordres_du_jour\\TXT\\page_' + str(compteur_page) + '.txt', fichier_PDF]
            
            pdf2txt.main(arg)                                                   #Convertir la page du PDF en texte
                
            with open('C:\\ContratsOuvertsMtl\\Ordres_du_jour\\TXT\\page_' + str(compteur_page) + '.txt', "r",) as f:
                reader = csv.reader(f, delimiter = "|")                         #Accéder au fichier texte généré

                for ligne in reader:                                            #Passer au travers du fichier texte généré
                
                    if est_dans_section_affaires_contractuelle == False:        #Indicateur si on est dans la section des contrats
                        if TITRE_SECTION_20 in str(ligne).encode("utf-8"):
                            est_dans_section_affaires_contractuelle = True
                
                    if TITRE_SECTION_30 in str(ligne).encode("utf-8"):          #Indicateur si on a fini de traiter la section des contrats
                        continuer = False
                        break
                    else:
                        if est_dans_section_affaires_contractuelle:             #Écrire la page dans le fichier fichier_TXT
                            if left(str(ligne),7) == "['Page ":                 #Ne pas écrire le numéro de page du pied-de-page
                                break
                            else:
                                #Ajouter la ligne dans le fichier fichier_TXT
                                fodj_traites.writerow(ligne)
                    
            f.close()       

    odj_traites.close()
            
    print      
    print("Fin du traitement") 

if __name__ == '__main__':
    main()
