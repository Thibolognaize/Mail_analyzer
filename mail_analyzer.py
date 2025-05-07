import pandas as pd
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Couleurs ANSI pour l'affichage
GREEN = '\033[92m'
YELLOW = '\033[93m'
RED = '\033[91m'
RESET = '\033[0m'

def afficher_message_accueil():
    """
    Affiche un message d'accueil pour l'utilisateur.
    """
    print(f"{YELLOW}Bienvenue dans le programme d'analyse d'emails !{RESET}")
    print(f"{YELLOW}Pensez à placer le fichier excel (.xlsx) à analyser dans le même dossier que ce script{RESET}")

def charger_fichier_excel():
    """
    Charge un fichier Excel à partir du nom fourni par l'utilisateur.
    Retourne un DataFrame pandas contenant les données du fichier.
    """
    while True:
        try:
            ExcelName = input("\nVeuillez entrer le nom du fichier: ")
            if ExcelName[-5:] != '.xlsx':
                ExcelName += '.xlsx'
            df = pd.read_excel(ExcelName)
            print(f"{GREEN}Chargement du fichier Excel {ExcelName}...{RESET}\n")
            return df
        except FileNotFoundError:
            print(f"{RED}Fichier introuvable. Veuillez vérifier le nom du fichier et réessayer.{RESET}")


def obtenir_colonne_email(df):
    """
    Demande à l'utilisateur de spécifier la colonne contenant les adresses email.
    Retourne le nom de la colonne sélectionnée.
    """
    while True:
        column_names = df.columns
        print(f"{YELLOW}Colonnes disponibles dans le fichier :{RESET}")
        print(column_names)
        email_column = input(f"{YELLOW}Entrez le nom de la colonne contenant les adresses email: {RESET}")
        if email_column in column_names:
            return email_column
        else:
            print(f"{RED}Nom de colonne invalide. Veuillez réessayer.{RESET}")

def analyser_emails(emails):
    """
    Analyse une série d'emails pour extraire les extensions, domaines, emails invalides et multiples.
    Retourne quatre listes : extensions, domaines, emails_invalides, emails_multiples.
    """
    extensions = []
    domaines = []
    emails_valides = []
    emails_invalides = []
    emails_multiples = []

    invalid_chars = [',', ';', ':', '!']
    valid_extensions = ['com', 'fr', 'net', 'org', 'mc', 'es', 'io', 'uk', 'be', 'eu', 'pro', 'no', 'pt', 'de', 'au', 'ch', 'art', 'bio', 'wine', 'solar', 'edu', 'pw']

    print(f"{GREEN}Analyse des adresses email en cours...{RESET}")

    for index, email_entry in enumerate(emails):
        email_entry = email_entry.strip()
        email = email_entry.split(';')[0]
        split_email = email_entry.split(';')

        if len(split_email) > 1:
            serie_email = ';'.join(split_email[1:]).strip()
            emails_multiples.append((index + 2, serie_email))

        if any(char in email for char in invalid_chars):
            emails_invalides.append((index + 2, email))
            continue

        # Vérification de la validité de l'email
        if '@' in email and '.' in email:
            # Vérification de l'extension (ex: .com, .fr, etc.)
            extension = email.split('.')[-1]
            if len(extension) >= 2 and len(extension) <= 7 and extension not in extensions:
                extensions.append(extension)

            # Vérification du domaine (ex: gmail, yahoo, etc.)
            domaine_part = email.split('@')[-1].split('.')
            if len(domaine_part) > 1:
                domaine = domaine_part[-2]
                if domaine not in domaines:
                    domaines.append(domaine)

            # Si l'extension est inconnue, on l'ajoute à la liste des emails invalides
            if extension in valid_extensions:
                emails_valides.append((index + 2, email))
            else:
                emails_invalides.append((index + 2, email))
        else:
            emails_invalides.append((index + 2, email))

    return extensions, domaines, emails_valides, emails_invalides, emails_multiples

def afficher_resultats(extensions, domaines, emails_valides, emails_invalides, emails_multiples):
    """
    Affiche les résultats de l'analyse des emails.
    """
    print(f"\n{GREEN}Analyse terminée. Voici les résultats :{RESET}")

    print(f"\n{YELLOW}Liste des extensions uniques :{RESET}")
    print(extensions)

    print(f"\n{YELLOW}Liste des domaines uniques :{RESET}")
    print(domaines)

    print(f"\n{YELLOW}Emails valides :{RESET}")
    for index, email in emails_valides:
        print(f"Index {index}: {email}")

    print(f"\n{YELLOW}Emails invalides :{RESET}")
    for index, email in emails_invalides:
        print(f"Index {index}: {email}")

    print(f"\n{YELLOW}Emails multiples :{RESET}")
    for index, email in emails_multiples:
        print(f"Index {index}: {email}")

def revise_file(df, emails_valides, emails_invalides, emails_multiples):
    """
    Corrige le fichier Excel en réécrivant les emails invalides et multiples.
    """
    invalid_chars = [',', ';', ':', '!']

    # Créer une copie du DataFrame pour appliquer les modifications
    df_corrected = df.copy()

    # Corriger les emails invalides
    for index, email in emails_invalides:
        corrected_email = email.lower()  # Mettre en minuscules
        corrected_email = corrected_email.replace(',', '.')  # Remplacer les virgules par des points
        corrected_email = ''.join(char for char in corrected_email if char not in invalid_chars)  # Retirer les caractères invalides

        # Mettre à jour le DataFrame avec l'email corrigé
        df_corrected.at[index - 2, df.columns[0]] = corrected_email

        # Ajouter l'email corrigé à la liste des emails valides
        emails_valides.append((index, corrected_email))

    # Corriger les emails multiples
    for index, email in emails_multiples:
        corrected_email = email.lower()  # Mettre en minuscules
        corrected_email = corrected_email.replace(',', '.')  # Remplacer les virgules par des points
        corrected_email = ''.join(char for char in corrected_email if char not in invalid_chars)  # Retirer les caractères invalides

        # Mettre à jour le DataFrame avec l'email corrigé
        df_corrected.at[index - 2, df.columns[0]] = corrected_email

        # Ajouter l'email corrigé à la liste des emails valides
        emails_valides.append((index, corrected_email))

    # Demander à l'utilisateur s'il souhaite enregistrer le fichier corrigé
    validation = input(f"\n{YELLOW}Voulez-vous enregistrer le fichier corrigé? (oui/non) : {RESET}")

    if validation.lower() == 'oui':
        # Enregistrer le nouveau fichier Excel
        output_file = "corrected_file.xlsx"
        df_corrected.to_excel(output_file, index=False)
        print(f"{GREEN}Fichier corrigé enregistré sous {output_file}.{RESET}")
    else:
        print(f"{RED}Enregistrement du fichier corrigé annulé.{RESET}")

    return df_corrected


def enregistrer_rapport(df, emails_invalides, emails_multiples):
    """
    Enregistre les résultats de l'analyse dans un fichier Excel et un fichier texte.
    """
    validation = input(f"\n{YELLOW}Voulez-vous enregistrer un nouveau fichier avec les données invalides surlignées? (oui/non) : {RESET}")

    if validation.lower() == 'oui':
        output_file = 'colored_data.xlsx'
        df.to_excel(output_file, index=False)

        book = load_workbook(output_file)
        ws = book.active
        fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

        all_emails_to_highlight = emails_invalides + emails_multiples
        for index, _ in all_emails_to_highlight:
            row = ws[index]
            for cell in row:
                cell.fill = fill

        book.save(output_file)
        print(f"{GREEN}Données enregistrées avec succès sous {output_file}.{RESET}")

        with open('rapport_emails.txt', 'w', encoding='utf-8') as file:
            file.write("Ce fichier contient un rapport des emails détectés invalides ou multiples. Avec 'Index' correspondant à la ligne Excel, et l'email associé à cette ligne. \n\n")
            file.write("Emails invalides :\n")
            for index, email in emails_invalides:
                file.write(f"Index {index}: {email}\n")
            file.write("\nEmails multiples :\n")
            for index, email in emails_multiples:
                file.write(f"Index {index}: {email}\n")

        print(f"{GREEN}Le rapport a été écrit dans le fichier rapport_emails.txt.{RESET}")
    else:
        print(f"{RED}Enregistrement annulé.{RESET}")

def main():
    """
    Fonction principale qui orchestre l'exécution du programme.
    """
    afficher_message_accueil()
    df = charger_fichier_excel()
    email_column = obtenir_colonne_email(df)
    emails = df[email_column]
    extensions, domaines, emails_valides, emails_invalides, emails_multiples = analyser_emails(emails)
    afficher_resultats(extensions, domaines, emails_valides ,emails_invalides, emails_multiples)
    revise_file(df, emails_valides,emails_invalides, emails_multiples)
    enregistrer_rapport(df, emails_invalides, emails_multiples)
    print(f"{GREEN}Fermeture du programme...{RESET}")
    time.sleep(2)

if __name__ == "__main__":
    main()
