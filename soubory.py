import os
from datetime import datetime
import win32com.client as win32 # potreba instalovat knihovnu -->> pip install pypiwin32

adresar = r'C:\Test' # zmenit cestu do adresare, ve kterem hledame soubory

validovane_spolecnosti = ['spolecnost_1','spolecnost_3'] # nazvy slozek spolecnosti,ktere jsou validovany
email = []

for i in validovane_spolecnosti:
     nazev_adresare = adresar + '\\' + i
     if os.path.exists(nazev_adresare) and os.path.isdir(nazev_adresare):
         obsah_adresare = os.listdir(nazev_adresare)
         for s in obsah_adresare:
             je_soubor = os.path.isfile(nazev_adresare + '\\' + s)
             if je_soubor == True:
                 if "a" in s: # do podminky vlozime vyraz,ktery hledame v nazvu souboru
                     vystup = (f'Dne {datetime.now().strftime("%d.%m.%Y %H:%M")} v adresari {nazev_adresare} jsme nasli soubor {s}')
                     print(vystup)
                     email.append(vystup)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jan.test@email.cz'  # spravna emailova adresa
#mail.Bcc = 'jana.testova@email.cz;jana.testova@seznam.cz'  // pokud bychom posilali skrytou kopii 
mail.Subject = 'Testovaci email z Pythonu' # zvolit predmet emailu 
mail.Body = "\n".join(email)
mail.Send()
         
#Dalsi rozsireni je sys.argv, ktery bude prebirat informaci od uzivatele,zda se ma odeslat email
