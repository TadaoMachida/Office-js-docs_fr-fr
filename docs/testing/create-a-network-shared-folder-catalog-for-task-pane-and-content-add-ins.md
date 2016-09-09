
# Chargement de version test des compléments Office

Vous pouvez installer un complément Office à des fins de test dans un client Office s’exécutant sur Windows à l’aide d’un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau. 

>**Remarque :** Pour tester un complément Office dans Office Online, voir [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md). Pour tester un complément sur un iPad ou un Mac, voir [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md ). Pour tester un complément Outlook, voir [Chargement de version test des compléments Outlook](sideload-outlook-add-ins-for-testing.md ).

Déployez uniquement le fichier manifeste vers le catalogue de dossiers partagés. Déployez l’application web en elle-même vers un serveur web et spécifiez l’URL dans l’élément **SourceLocation** du fichier manifeste.

 >**Important :**  Pour contribuer à sécuriser les compléments accédant à des services et données externes, votre application doit utiliser un protocole sécurisé tel que HTTPS (Hypertext Transfer Protocol Secure) pour se connecter aux services et données externes. Vous devez utiliser HTTPS si votre complément utilise des commandes de complément.

## Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

2. Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.

3. Ouvrez l’onglet **Partage**.

4. Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. 

5. Choisissez **Partager** > **Terminer** > **Fermer**.

## Spécifier le dossier partagé en tant que catalogue approuvé

      
3. Ouvrez un nouveau document dans Excel, Word ou PowerPoint.
    
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
5. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
6. Choisissez **Catalogues de compléments approuvés**.
    
7. Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.
    
8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

9. Fermez l’application Office afin que vos modifications prennent effet.
    
## Charger votre complément


1. Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés.

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.


## Ressources supplémentaires

- [Utilisation de la journalisation runtime pour déboguer votre manifeste](../develop/use-runtime-logging-to-debug-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
    
