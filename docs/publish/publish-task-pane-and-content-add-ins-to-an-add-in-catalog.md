
# Publier des compléments de contenu et du volet Office dans un catalogue de compléments sur SharePoint

Un catalogue de compléments est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour les utiliser dans leur organisation. Lorsqu’un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.

>**Remarque :** Les catalogues de compléments sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud VersionOverrides du [manifeste de complément](../overview/add-in-manifests.md).

Les catalogues SharePoint ne sont pas pris en charge dans Office 2016 pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à l’[Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).   

## Configuration d’un catalogue de compléments sur SharePoint

1. Accédez au **site Administration centrale** (**Démarrer** > **Tous les programmes** > **Produits Microsoft SharePoint 2013** > **Administration centrale SharePoint 2013**).
    
2. Dans le volet Office de gauche, cliquez sur  **Compléments**.
    
3. Sur la page  **Compléments**, sous  **Gestion des compléments**, choisissez  **Gérer le catalogue de compléments**.
    
4. Sur la page  **Gérer le catalogue de compléments**, vérifiez que vous avez sélectionné l’application web appropriée dans  **Sélecteur d’applications web**.
    
5. Choisissez  **Afficher les paramètres du site**.
    
6. Sur la page  **Paramètre du site**, choisissez  **Administrateurs de collections de sites** pour spécifier les administrateurs de collection de sites, puis choisissez **OK**.
    
7. Pour accorder des autorisations de site aux utilisateurs, choisissez  **Autorisations de site**, puis choisissez  **Accorder des autorisations**.
    
8. Dans la boîte de dialogue  **Partager le site de catalogue d’applications**, spécifiez des utilisateurs de site, définissez les autorisations appropriées pour ces derniers, puis éventuellement d’autres options, puis choisissez  **Partager**.
    
9. Pour ajouter des compléments au catalogue de compléments Office, choisissez **Compléments Office**.

## Configuration d’un catalogue de compléments sur Office 365

1. Sur la page Centre d’administration Office 365, sélectionnez **Administrateur**, puis **SharePoint**.
    
2. Dans le volet Office situé à gauche, cliquez sur  **Compléments**.
    
3. Sur la page  **Compléments**, cliquez sur  **Catalogue de compléments**.
    
4. Sur la page  **Site de catalogue de compléments**, cliquez sur  **OK** pour accepter l’option par défaut et créer un site de catalogue de compléments.
    
5. Sur la page  **Créer une collection de sites de catalogue de compléments**, indiquez le titre de votre site de catalogue de compléments.
    
6. Spécifiez l’adresse du site web.
    
7. Définissez l’option  **Quota de stockage** sur la plus faible valeur possible (actuellement 110). Vous n’installerez que des packages de complément sur cette collection de sites et ils sont peu volumineux.
    
8. Définissez l’option  **Quota de ressources du serveur** sur 0 (zéro). (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue de compléments.)
    
9. Sélectionnez  **OK**.
    
Pour ajouter des complément au site de catalogue de compléments, accédez au site que vous venez de créer. Dans le volet de navigation de gauche, choisissez  **Compléments Office**, puis, pour télécharger un fichier manifeste Office, sélectionnez  **Nouveau complément**.    

## Publication dans un catalogue de compléments


1. Accédez au catalogue de compléments :

    1- Ouvrez la page principale de l’Administration centrale de SharePoint.
    
    2- Sélectionnez **Compléments**.
    
    3- Sélectionnez **Gérer le catalogue de compléments**.
    
    4- Sélectionnez le lien fourni, puis choisissez **Compléments Office** dans la barre de navigation située à gauche.
    
2. Sélectionnez le lien **Cliquer pour ajouter un nouvel élément**.
    
3. Choisissez **Parcourir**, puis spécifiez le [manifeste](../../docs/overview/add-in-manifests.md) à télécharger.
    
    Les compléments de contenu et de volet Office de ce catalogue sont désormais disponibles dans la boîte de dialogue **Compléments Office**. Pour y accéder, choisissez **Mes compléments** sous l’onglet **Insérer**, puis choisissez **MON ORGANISATION**.
    
Une fois les manifestes de compléments chargés dans le catalogue de compléments Office, les utilisateurs peuvent accéder aux compléments en procédant comme suit :


1. Dans l’application Office, accédez à **Fichier**  >  **Options**  >  **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité**  >  **Catalogues de compléments approuvés**.
    
2. Spécifiez l’URL de la  _collection de sites SharePoint parente_ du catalogue de compléments. Par exemple, si l’URL du catalogue de compléments Office est :
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    Spécifiez simplement l’URL de la collection de sites parente :
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. Fermez puis rouvrez l’application Office. Le catalogue de compléments est disponible dans la boîte de dialogue **Compléments Office**.
    
Par ailleurs, un administrateur peut spécifier un catalogue de compléments Office sur SharePoint à l’aide de la stratégie de groupe. Pour plus d’informations, voir la section « Utilisation de la stratégie de groupe pour gérer la façon dont les utilisateurs peuvent installer et utiliser les compléments Office » dans la rubrique [Vue d’ensemble des compléments Office](https://technet.microsoft.com/en-us/library/jj219429.aspx) sur TechNet.

