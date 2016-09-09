
# Création d’un catalogue de dossiers partagés réseau pour les compléments de contenu ou du volet Office


Un catalogue de dossiers partagés permet de publier les manifestes de Compléments Office de volet de tâches et de contenu sur un partage de fichiers réseau. Les utilisateurs peuvent alors acquérir des compléments en spécifiant ce partage de fichiers comme catalogue approuvé à l’aide des étapes de la procédure suivante.

Le fichier manifeste est un fichier XML qui vous permet de décrire de façon déclarative la manière dont votre complément doit être activé lorsqu’un utilisateur final l’installe et l’utilise avec des documents et des applications Office. Pour plus d’informations, voir [Manifeste XML des compléments Office](../../docs/overview/add-in-manifests.md).

Le fichier manifeste est le seul fichier que vous devez déployer vers le catalogue de dossiers partagés. Déployez l’application web elle-même vers un serveur web, et indiquez l’URL dans l’élément  **SourceLocation** du fichier manifeste.

 >**Important :**  Pour contribuer à sécuriser les compléments accédant à des services et données externes, votre application doit utiliser un protocole sécurisé tel que HTTPS (Hypertext Transfer Protocol Secure) pour se connecter aux services et données externes. Vous devez utiliser HTTPS si votre complément utilise des commandes de complément.


## Spécification d’un partage de fichiers comme catalogue approuvé


1. Créez un dossier sur un partage réseau, par exemple  `\\MyShare\MyManifests`.
    
2. Placez dans ce partage de fichiers les fichiers manifestes pour les compléments du volet Office et de contenu que vous souhaitez publier.
    
3. Ouvrez un nouveau document dans Excel, Word ou PowerPoint.
    
4. Choisissez l’onglet  **Fichier**, puis choisissez  **Options**.
    
5. Choisissez  **Centre de gestion de la confidentialité**, puis cliquez sur le bouton  **Paramètres du Centre de gestion de la confidentialité**.
    
6. Choisissez  **Catalogues de compléments approuvés**.
    
7. Dans la zone  **URL du catalogue**, entrez le chemin d’accès du partage réseau que vous avez créé à l’étape 1, puis choisissez  **Ajouter un catalogue**.
    
8. Activez la case à cocher  **Afficher dans le menu**, puis choisissez  **OK**.
    
Après avoir exécuté ces étapes, vous pouvez sélectionner  **Mes compléments** dans l’onglet **Insérer** du ruban, puis choisir **Dossier partagé** en haut de la boîte de dialogue **Compléments Office** pour insérer un complément de volet de tâches ou de contenu de ce catalogue.

Tous les autres fichiers manifestes que vous placerez dans ce partage de fichiers seront accessibles aux utilisateurs qui ont spécifié ce catalogue de dossiers partagés.


## Ressources supplémentaires



- [Publier votre complément Office](../publish/publish.md)
    

