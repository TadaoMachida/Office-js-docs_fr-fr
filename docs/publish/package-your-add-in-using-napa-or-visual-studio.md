
# Empaquetage de votre complément à l’aide de Napa ou de Visual Studio pour préparer la publication

Votre package de complément Office contient un fichier XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers de l’application web de votre projet séparément.

## Empaquetage d’un Complément Office créé à l’aide de Outils de développement Office 365 « Napa »



1. Dans Napa, sur le côté de la page, cliquez sur le bouton  **Publier** ( ![Bouton Publier](../../images/Apps_NAPA_Publish.png)).
    
2. Dans la boîte de dialogue  **Paramètres de publication**, cliquez sur le bouton  **Suivant**.
    
3. Indiquez l’URL du site web qui hébergera les fichiers de contenu de votre complément (par exemple, les fichiers JavaScript et HTML par défaut de votre projet), puis cliquez sur le bouton **Publier**.
    
4. Dans la boîte de dialogue  **Publication réussie**, cliquez sur le lien  **Emplacement de publication**.
    
    Une bibliothèque de documents contenant le fichier manifeste XML de votre complément et les fichiers de contenu web s’affiche. 
    
Ensuite, copiez manuellement les fichiers de contenu web (feuilles de style, fichiers JavaScript et fichiers HTML) sur le serveur web qui héberge le site web que vous avez indiqué dans la boîte de dialogue  **Paramètres de publication**.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). 


## Déploiement de votre projet web et empaquetage de votre complément à l’aide de Visual Studio 2015



### Pour déployer votre projet Web


1. Dans l’ **Explorateur de solutions**, ouvrez le menu contextuel du projet d’complément, puis sélectionnez  **Publier**.
    
    La page **Publier votre complément** s’ouvre.
    
2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau …** pour créer un profil.
    
     >**Remarque**  Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez  **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.
    
    Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page **Publier votre complément**, cliquez sur le lien **Déployer votre projet Web**.
    
    La boîte de dialogue **Publier le site web** s’affiche. Pour plus d’informations sur l’utilisation de cet Assistant, reportez-vous à l’article relatif à la [procédure de déploiement d’un projet web à l’aide de la publication en un clic dans Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### Empaquetage de votre complément


1. Sur la page  **Publier votre complément**, cliquez sur le lien  **Empaqueter le complément**.
    
    L’Assistant **Publication des compléments SharePoint et Office** apparaît.
    
2. Dans la liste déroulante  **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur  **Terminer**.
    
    Vous devez spécifier une adresse qui commence par le préfixe HTTPS pour l’Assistant. En règle générale, l’utilisation d’un point de terminaison HTTPS pour votre site web est la meilleure approche, mais cela n’est pas obligatoire si vous ne comptez pas publier votre complément sur l’Office Store. Une fois le package créé, vous pouvez ouvrir le manifeste dans le Bloc-notes et remplacer le préfixe HTTPS de votre site web par le préfixe HTTP. Pour plus d’informations, voir [Pourquoi mes compléments doivent-ils être sécurisés par une protection SSL ?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**Remarque**  Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication. 
    
Si vous prévoyez de soumettre votre complément à l’Office Store, vous pouvez cliquer sur le lien **Effectuer un test de validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez régler tous ces problèmes avant de soumettre votre complément au magasin.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans  `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## Ressources supplémentaires



- [Publier votre complément Office](../publish/publish.md)
    
- [Soumission des compléments SharePoint et Office, ainsi que des applications web Office 365 dans l’Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
