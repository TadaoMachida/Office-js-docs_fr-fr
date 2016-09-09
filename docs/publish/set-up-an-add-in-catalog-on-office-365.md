
# Configurer un catalogue de compléments dans Office 365

Un catalogue de complément est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des Compléments SharePoint et des Compléments Office. Les administrateurs peuvent télécharger des fichiers manifeste de Compléments Office vers le catalogue de compléments pour les utiliser dans leur organisation. Quand un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé (en configurant la stratégie de groupe ou en indiquant le catalogue approuvé dans l’onglet  **Catalogues de compléments approuvés** de la boîte de dialogue **Options**, et en sélectionnant  **Fichier** > **Options** > **Centre de gestion de la confidentialité** > **Paramètres du Centre de gestion de la confidentialité** > **Catalogues de compléments approuvés**), les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.

## Configuration d’un catalogue de complément sur SharePoint Online


1. Sur la page Centre d’administration Office 365, sélectionnez  **Administrateur**, puis  **SharePoint**.
    
2. Dans le volet Office situé à gauche, cliquez sur  **Compléments**.
    
3. Sur la page  **Compléments**, cliquez sur  **Catalogue de compléments**.
    
4. Sur la page  **Site de catalogue de compléments**, cliquez sur  **OK** pour accepter l’option par défaut et créer un site de catalogue de compléments.
    
5. Sur la page  **Créer une collection de sites de catalogue de compléments**, indiquez le titre de votre site de catalogue de compléments.
    
6. Spécifiez l’adresse du site web.
    
7. Définissez l’option  **Quota de stockage** sur la plus faible valeur possible (actuellement 110). Vous n’installerez que des packages de complément sur cette collection de sites et ils sont peu volumineux.
    
8. Définissez l’option  **Quota de ressources du serveur** sur 0 (zéro). (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue de compléments.)
    
9. Sélectionnez  **OK**.
    
Pour ajouter des complément au site de catalogue de compléments, accédez au site que vous venez de créer. Dans le volet de navigation de gauche, choisissez  **Compléments Office**, puis, pour télécharger un fichier manifeste Office, sélectionnez  **Nouveau complément**.


## Ressources supplémentaires


- [Publier des compléments dans un catalogue de compléments](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)

    

