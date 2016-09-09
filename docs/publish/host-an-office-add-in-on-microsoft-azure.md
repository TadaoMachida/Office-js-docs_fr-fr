
# Héberger un complément pour Office sur Microsoft Azure

Le complément Office le plus simple est constitué d’un fichier manifeste XML et d’une page HTML. Le fichier manifeste XML décrit les caractéristiques du complément, telles que son nom, les applications clientes Office dans lesquelles il peut fonctionner et l’URL de la page HTML du complément. La page HTML est contenue sur le site web d’un complément Office et les utilisateurs peuvent la voir et interagir avec elle lors de l’installation et de l’exécution de votre complément. 

Vous pouvez héberger le site web d’un complément Office sur plusieurs plateformes d’hébergement web, y compris Azure. Pour héberger un complément Office sur Azure, publiez-le sur un site web Azure. 

Cette rubrique suppose que vous n’avez jamais utilisé Azure. À la fin, vous disposerez d’un complément Office simple dont le site web est hébergé sur Azure. Vous apprendrez à :

- ajouter un catalogue de compléments approuvés à Office 2013 ;
    
- à créer un site web dans Azure à l’aide de Visual Studio 2015 ou du portail de gestion Azure
    
- à publier et héberger un complément Office sur un site web Azure
    

**Site web d’un complément Office hébergé sur Azure**


![Site web d’un complément Office hébergé dans Microsoft Azure](../../images/off15app_HowToPublishA4OtoAzure_fig17.png)


## Configurer votre ordinateur de développement avec le kit de développement logiciel Azure SDK pour .NET, un abonnement Azure et Office 2013



1. Installez le kit de développement logiciel Azure SDK pour .NET à partir de la [page Téléchargements d’Azure](http://azure.microsoft.com/en-us/downloads/). Si Visual Studio n’est pas installé, Visual Studio Express pour le web est installé avec le kit de développement logiciel.
    
    - Sous  **Langues**, choisissez  **.NET**.
    
    - Choisissez la version du kit de développement logiciel Azure SDK pour .NET correspondant à votre version de Visual Studio, si Visual Studio est déjà installé.
    
    - Quand on vous demande si vous souhaitez exécuter ou enregistrer l’exécutable d’installation, choisissez  **Exécuter**.
    
    - Dans la fenêtre Web Platform Installer, choisissez  **Installer**.
    
2. Installez Office 2013 si ce n’est pas déjà fait. 
    
     >**Remarque :**  Vous pouvez obtenir une [version d’évaluation pour un mois](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).
3. Accédez à votre compte Azure.
    
     >**Remarque :**  Si vous êtes abonné à MSDN, [vous bénéficiez d’un abonnement à Azure dans le cadre de votre abonnement à MSDN](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/). Si vous n’êtes pas abonné à MSDN, vous pouvez toujours [obtenir une version d’évaluation gratuite d’Azure sur le site web Windows Azure](https://azure.microsoft.com/en-us/pricing/free-trial/). 

Pour que la procédure reste simple et centrée sur l’utilisation d’Azure avec un complément Office, utilisez un partage de fichiers local en tant que catalogue approuvé dans lequel vous stockez le fichier manifeste XML du complément. Pour un complément que vous avez l’intention d’utiliser dans une ou plusieurs entreprises, vous pouvez conserver le fichier manifeste du complément dans SharePoint ou publier le complément dans l’Office Store. 


## Étape 1 : créer un partage de fichiers réseau pour héberger votre fichier manifeste de complément



1. Ouvrez l’Explorateur de fichiers (ou l’Explorateur Windows si vous utilisez Windows 7 ou une version antérieure de Windows) sur votre ordinateur de développement.
    
2. Cliquez avec le bouton droit sur le lecteur C:\, puis choisissez **Nouveau** > **Dossier**.
    
3. Nommez le nouveau dossier AddinManifests.
    
4. Cliquez avec le bouton droit sur le dossier AddinManifests, puis choisissez  **Partager avec**  >  **Des personnes spécifiques**.
    
5. Dans  **Partage de fichiers**, sélectionnez la flèche déroulante vers le bas, puis choisissez  **Tout le monde**  >  **Ajouter**  >  **Partager**.
    

## Étape 2 : ajouter le partage de fichiers au catalogue de compléments approuvés de sorte que les applications clientes Office approuvent l’emplacement où vous installez les compléments Office



1.  Démarrez Word 2013 et créez un document. (Même si nous utilisons Word 2013 dans cet exemple, vous pourriez utiliser n’importe quelle application Office qui prend en charge les compléments Office comme Excel, Outlook, PowerPoint ou Project 2013.)
    
2.  Choisissez **Fichier**  >  **Options**.
    
3.  Dans **Options Word**, choisissez  **Centre de gestion de la confidentialité**, puis  **Paramètres du Centre de gestion de la confidentialité**. 
    
4.  Dans le **Centre de gestion de la confidentialité**, cliquez sur  **Catalogues de compléments approuvés**. Saisissez le chemin d’accès UNC (Universal Naming Convention) pour le partage de fichiers que vous avez créé précédemment en tant qu’**URL du catalogue**. Par exemple,  \\YourMachineName\AddinManifests. Ensuite, choisissez  **Ajouter un catalogue**. 
    
5. Cochez la case correspondant à  **Afficher dans le menu**. Lorsque vous stockez un fichier manifeste XML de complément sur un partage qui est un catalogue de compléments approuvés, le complément apparaît sous  **Dossier partagé** dans la boîte de dialogue **Compléments Office**.
    

## Étape 3 : créer un site web dans Azure


Il existe plusieurs façons de créer un site web Azure vide. Si vous utilisez Visual Studio 2015, suivez les étapes dans la section [Utilisation de Visual Studio 2015](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2015) pour créer un site web Azure à partir de l’IDE de Visual Studio. Vous pouvez également suivre les étapes dans la section [Utilisation du portail de gestion Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-management-portal) pour créer le site web Azure.


### Utilisation de Visual Studio 2015



1. Dans Visual Studio, dans le menu  **Afficher**, sélectionnez **Explorateur de serveurs**. Cliquez avec le bouton droit de la souris sur  **Azure** et choisissez **Se connecter à un abonnement Microsoft Azure**. Suivez les instructions pour vous connecter à votre abonnement Azure.
    
2. Dans Visual Studio, dans  **Explorateur de serveurs**, développez  **Azure**, cliquez avec le bouton droit de la souris sur **App Service**, puis choisissez  **Créer une application web**.
    
3. Dans la boîte de dialogue  **Créer une application web sur Windows Azure**, fournissez les informations suivantes :
    
      - Entrez le **Nom de l’application web** unique pour votre site. Azure vérifie que le nom de site est unique dans le domaine azurewebsites.net.
    
  - Sélectionnez le plan **App Service** utilisé pour autoriser la création de ce site web. Si vous créez un plan, vous devez aussi le nommer.
    
  - Sélectionnez le **Groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.
    
  - Choisissez une **Région** géographique qui vous convient.
    
  - Pour le champ  **Serveur de base de données :**, acceptez la valeur par défaut  **Aucune base de données**, puis choisissez **Créer**.
    

    Le nouveau site web s’affiche sous le groupe de ressources choisi sous **Service d'application** sous **Azure** dans **Explorateur de serveurs**.
    
4. Cliquez avec le bouton droit sur le nouveau site web, puis choisissez  **Afficher dans le navigateur**. Votre navigateur s’ouvre et affiche une page web avec le message « Ce site web a été créé ».
    
5. Dans la barre d’adresse du navigateur, modifiez l’URL du site web afin qu’il utilise le protocole HTTPS, puis appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé. Le modèle de complément Office exige que les compléments utilisent le protocole HTTPS.
    
6. Dans Visual Studio 2015, cliquez avec le bouton droit sur le nouveau site web dans l’**Explorateur de serveurs**, choisissez  **Télécharger le profil de publication**, puis enregistrez le profil sur votre ordinateur. Le profil de publication contient vos informations d’identification et vous permet de passer à l’[Étape 5 : publier le complément Office sur le site web Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-5-publish-your-office-add-in-to-the-azure-website).
    

### Utilisation du portail de gestion Azure



1. Connectez-vous au [portail de gestion Azure](https://manage.windowsazure.com/) avec votre compte Azure.
    
2. Choisissez  **NOUVEAU**  >  **CALCULER**  >  **APPLICATION WEB**  >  **CRÉATION RAPIDE**. 
    
3. Sous  **URL**, entrez un nom de site unique pour compléter l’URL du site web. Le portail de gestion vérifie que le nom du site est unique dans le domaine azurewebsites.net.
    
4. Choisissez une **RÉGION** géographique appropriée pour votre site.
    
5. Choisissez  **CRÉER UNE APPLICATION WEB**. Le portail de gestion Azure crée le site web et vous redirige vers la page des  **sites web** où vous pouvez voir l’état du site web.
    
    Lorsque l’état du site web est **En cours d’exécution**, choisissez l’URL du site web sous la colonne **NOM**. Votre navigateur s’ouvre et affiche une page web avec le message **Votre application web a été créée !**. 
    
    Dans la barre d’adresse du navigateur, modifiez l’URL du site web afin qu’il utilise le protocole HTTPS, puis appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé. Le modèle de complément Office exige que les compléments utilisent le protocole HTTPS.
    
6. Sur la page des  **applications web**, cliquez sur le nouveau site web.
    
7. Sous  **Publier votre application**, choisissez  **Télécharger le profil de publication**, puis enregistrez le profil de publication sur votre ordinateur. N’oubliez pas le nom de fichier et l’emplacement, car vous en aurez besoin plus tard.
    
    Le profil de publication contient vos informations d’identification et vous permet de publier en toute sécurité sur Azure. 
    

## Étape 4 : créer un complément Office dans Visual Studio.



1. Démarrez Visual Studio en tant qu’administrateur.
    
2. Choisissez  **Fichier**  >  **Nouveau**  >  **Projet**.
    
3. Sous  **Modèles**, développez  **Visual C#** (ou **Visual Basic**), développez  **Office/SharePoint** et choisissez  **Compléments Office**.
    
4. Choisissez  **Complément Office**, puis cliquez sur **OK** pour accepter les paramètres par défaut.
    
5. Quand le message  **Créer un complément Office** apparaît, conservez le choix par défaut pour un complément du volet Office puis cliquez sur **Suivant**.
    
6. Sur la page suivante, désélectionnez toutes les cases, sauf pour Word, puis sélectionnez  **Terminer**.
    
Votre complément Office de base est créé et prêt à être publié sur Azure. Dans la mesure où nous vous expliquons comment publier sur Azure, aucune modification ne devra être apportée à l’exemple de complément que vous avez créé avec le modèle de complément Office standard dans Visual Studio.

## Étape 5: publier le complément Office sur le site web d’Azure



1. Avec votre exemple de complément ouvert dans Visual Studio, développez le nœud de solution dans l’**Explorateur de solutions** pour voir les deux projets de la solution.
    
2. Cliquez avec le bouton droit sur le projet web, puis choisissez  **Publier**. 
    
    Le projet web contient les fichiers de site web du complément Office, et il s’agit donc du projet que vous publiez sur Azure.
    
3. Dans  **Publier le site Web**, choisissez  **Importer**. 
    
4. Dans  **Importer les paramètres de publication**, sélectionnez  **Parcourir**, puis recherchez l’emplacement où vous avez enregistré votre profil de publication précédemment dans cette rubrique. Sélectionnez  **OK** pour importer votre profil.
    
5. Dans  **Publier le site web**, dans l’onglet  **Connexion**, acceptez les valeurs par défaut puis cliquez sur **Suivant**. 
    
    Choisissez **Suivant ** à nouveau pour accepter les paramètres par défaut.
    
6. Dans l’onglet  **Aperçu**, choisissez  **Démarrer l’aperçu**. L’aperçu vous montre tous les fichiers du projet web qui seront publiés sur le site web d’Azure.
    
7. Cliquez sur  **Publier**. Visual Studio publie le projet web pour votre complément Office sur votre site web Azure. 
    
8. Quand Visual Studio termine la publication du projet web, votre navigateur s’ouvre et affiche une page web avec le texte « Cette application web a été créée ». Ceci est la page par défaut actuelle du site web.
    
    Pour afficher la page web de votre complément, modifiez votre l’URL de manière à utiliser le protocole https: et ajoutez le chemin d’accès de la page HTML par défaut de votre complément. Par exemple, l’URL modifiée doit être du type https://VotreDomaine.azurewebsites.net/Addin/Home/Home.html. Cela permet de confirmer que de site web de votre complément est hébergé sur Azure. Copiez cette URL, vous en aurez besoin lorsque vous modifierez le fichier manifeste du complément plus loin dans cette rubrique.
    

## Étape 6 : modifier le fichier manifeste du complément pour qu’il pointe vers le complément Office sur Azure



1. Dans Visual Studio avec l’exemple de complément Office ouvert dans l’**Explorateur de solutions**, développez la solution pour que les deux projets s’affichent.
    
2. Développez le projet de complément Office, par exemple **OfficeAdd-in1**, cliquez avec le bouton droit sur le dossier de manifeste, puis sur  **Ouvrir**. La page de propriétés du manifeste du complément apparaît.
    
3. Pour le champ  **Emplacement source :**, saisissez l’URL de la page HTML principale du complément que vous avez copiée à l’étape précédente après avoir publié le complément. Par exemple, https://YourDomain.azurewebsites.net/Addin/Home/Home.html. 
    
4. Choisissez  **Fichier**, puis  **Enregistrer tout**. Fermez la page de propriétés du manifeste du complément.
    
5. Retournez dans l’**Explorateur de solutions**, cliquez avec le bouton droit sur le dossier de manifeste et choisissez  **Ouvrir le dossier dans l’Explorateur de fichiers**.
    
6. Copiez le fichier manifeste du complément, par exemple OfficeAdd-in1.xml. 
    
7. Accédez au partage de fichiers réseau que vous avez créé plus tôt dans la rubrique et collez le fichier manifeste dans le dossier.
    

## Étape 7 : insérer et exécuter le complément dans l’application cliente Office



1. Démarrez Word et ouvrez un nouveau document.
    
2. Sur le ruban, choisissez  **Insérer**  >  **Mes applications**, puis  **Voir tout**.
    
3. Dans la boîte de dialogue  **Applications pour Office**, choisissez **DOSSIER PARTAGÉ**. Les applications clientes Office qui fonctionnent avec le modèle des compléments Office analysent le dossier indiqué comme catalogue de compléments approuvés et affichent les compléments dans la boîte de dialogue. Vous devez voir l’icône pour votre exemple de complément.
    
4. Choisissez l’icône de votre complément, puis choisissez **Insérer**. Le complément est inséré sur le côté de l’application cliente.
    
5. Vérifiez que le complément fonctionne en créant un texte dans le document, puis en sélectionnant ce texte et en choisissant  **Obtenir des données à partir de la sélection**.
    

## Ressources supplémentaires



- [Publier votre complément Office](../publish/publish.md)
    
- [Empaquetage de votre complément à l’aide de Napa ou de Visual Studio pour préparer la publication](../publish/package-your-add-in-using-napa-or-visual-studio.md)
    
