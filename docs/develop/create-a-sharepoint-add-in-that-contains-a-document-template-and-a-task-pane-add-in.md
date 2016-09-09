
# Créer un complément SharePoint qui contient un modèle de document et un complément du volet Office


Vous pouvez créer un SharePoint Add-in qui inclut un modèle de document (par exemple, une note de frais). Le document peut inclure un complément du volet Office qui interagit avec des données SharePoint. Par exemple, les utilisateurs peuvent remplir les champs d’une facture à l’aide des données Business Connectivity Services (BCS) ou créer une note de frais en sélectionnant une catégorie de frais dans une liste SharePoint.

Cette procédure pas à pas vous montre comment créer un SharePoint Add-in qui inclut un classeur Excel. Le classeur Excel contient un complément du volet Office qui utilise l’interface REST fournie par SharePoint 2013 pour remplir une zone de liste déroulante avec des données SharePoint dans le complément du volet Office.


## Conditions préalables


Installez les composants suivants avant de commencer :




- Environnement de développement SharePoint :
    
      - To develop SharePoint Add-ins that target SharePoint in Office 365, see [How to: Set up an environment for developing SharePoint Add-ins on Office 365](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29).
    
  - Pour développer des compléments SharePoint qui ciblent une installation locale de SharePoint, reportez-vous à [la procédure : Configurer un environnement de développement local pour les compléments SharePoint](http://msdn.microsoft.com/en-us/library/office/apps/fp179923%28v=office.15%29).
    
- [Visual Studio 2015 et outils de développement Microsoft Office](https://www.visualstudio.com/features/office-tools-vs)
    
- Excel 2013 ou un compte Office 365.
    

## Créer un projet de SharePoint Add-in dans Visual Studio



1. Démarrez Visual Studio.
    
2. Dans la barre de menus, choisissez **Fichier**, **Nouveau**, **Projet**.
    
    La boîte de dialogue **Nouveau projet** s’ouvre.
    
3. Dans le volet des modèles, sous le nœud de la langue utilisée, développez **Office/SharePoint**, puis choisissez **Compléments Office**.
    
4. Dans la liste des types de projet, choisissez **Complément SharePoint**, nommez le projet OfficeEnabledAddin, puis cliquez sur le bouton **OK**.
    
    La boîte de dialogue **Nouveau complément SharePoint** s’affiche.
    
5. Dans la liste déroulante **Quel site SharePoint souhaitez-vous utiliser pour déboguer votre complément ?**, choisissez ou entrez l’URL d’un site SharePoint.
    
6. Dans la liste déroulante **Comment souhaitez-vous héberger votre complément SharePoint ?**, choisissez **Hébergement par SharePoint**, puis cliquez sur **Suivant**.
    
     >**Remarque** : ce scénario fonctionne uniquement avec les options hébergées par SharePoint ou par le fournisseur figurant dans la liste déroulante **Comment souhaitez-vous héberger votre complément SharePoint ?**.
7. Sur la page suivante, sélectionnez **SharePoint 2013**, puis cliquez sur le bouton **Terminer** pour fermer la boîte de dialogue.
    

## Ajouter un élément de complément du volet Office


Ajoutez ensuite un Complément Office au projet. Vous pouvez ajouter n’importe quel type de complément. Pour cette procédure pas à pas, nous allons ajouter un complément du volet Office.


1. Dans l’**Explorateur de solutions**, choisissez le nœud de projet **OfficeEnabledAddin**.
    
2. Dans le menu **Projet**, choisissez **Ajouter un nouvel élément**.
    
3. Dans la boîte de dialogue **Ajouter un nouvel élément**, sélectionnez **Office/SharePoint**, puis cliquez sur **Complément Office**.
    
4. Nommez le complément du volet Office MyTaskPaneAddin, puis cliquez sur le bouton **Ajouter**.
    
    La boîte de dialogue **Créer un complément pour Office** s’affiche.
    
5. Dans la boîte de dialogue **Créer un complément pour Office**, sélectionnez **Volet Office**, puis **Suivant**. Sur la page suivante, désactivez les cases à cocher **Word** et **PowerPoint**, puis sélectionnez **Suivant**.
    
6. Dans la page **Voulez-vous que votre complément Office s’affiche dans un nouveau document ou un document existant ?**, choisissez **Créer un document et insérer mon complément**, puis cliquez sur le bouton **Terminer**.
    
    Visual Studio ajoute une bibliothèque de documents et un modèle de classeur pour la bibliothèque. Le classeur contient un complément de volet Office.
    

## Ajouter une bibliothèque de documents


Dans cette procédure, vous allez ajouter une bibliothèque de documents et faire du classeur le modèle par défaut de la bibliothèque de documents.


1. Dans l’**Explorateur de solutions**, choisissez le nœud de projet **OfficeEnabledAddin**.
    
2. Dans le menu **Projet**, choisissez **Ajouter un nouvel élément**.
    
3. Dans la boîte de dialogue **Ajouter un nouvel élément**, sélectionnez **Office/SharePoint**, puis **Liste**. Nommez la liste MyDocumentLibrary, puis cliquez sur le bouton **Ajouter**.
    
4. Dans l’**Assistant Personnalisation de SharePoint**, sélectionnez l’option **Créer un modèle de liste personnalisé et une instance de liste correspondante**.
    
5. Dans la liste déroulante située sous cette option, sélectionnez **Bibliothèque de documents**, puis cliquez sur le bouton **Suivant**.
    
6. Dans la page **Choisissez un modèle pour cette bibliothèque de documents. Les documents créés par les utilisateurs dans cette bibliothèque reposeront sur ce modèle.**, choisissez **Utilisez le document suivant comme modèle de cette bibliothèque**, puis choisissez le bouton **Parcourir**.
    
7. Dans la boîte de dialogue **Ouvrir**, ouvrez le dossier **OfficeDocuments**, sélectionnez le fichier **MyTaskPaneApp.xlsx**, choisissez le bouton **Ouvrir**, choisissez le bouton **Terminer**, puis fermez le concepteur de liste.
    
8. Dans l’**Explorateur de solutions**, choisissez le nœud de projet **OfficeEnabledAddin**.
    
9. Dans le menu **Affichage**, choisissez **Fenêtre Propriétés**.
    
10. Dans l’**Explorateur de solutions**, choisissez le fichier **AppManifest.xml**.
    
11. Choisissez **Affichage**, **Concepteur**.
    
12. Dans le concepteur de manifeste, définissez la valeur de la **page de démarrage** sur ~appWebUrl/Lists/MyDocumentLibrary. Cela la convertit en une valeur de OfficeEnabledAddin/Lists/MyDocumentLibrary.
    
     >**Remarque** : cette URL fait référence à la bibliothèque de documents. Vous devez utiliser le jeton ~appWebUrl au début des URL dans votre manifeste de complément Office faisant référence à des éléments dans le complément web. Pour plus d’informations sur les jetons d’URL dans un projet de complément SharePoint, consultez l’article [Chaînes URL et jetons dans les compléments pour SharePoint](http://msdn.microsoft.com/library/800ec8cd-a448-46bc-b41e-d4030eeb4048%28Office.15%29.aspx).
13. Fermez le concepteur de manifeste pour enregistrer la modification.
    

## Utiliser des données SharePoint dans le volet de tâches


Dans cette procédure, vous allez afficher la liste des utilisateurs du site à l’aide de l’interface REST fournie par SharePoint 2013.

Dans cet exemple, seules les données de liste SharePoint sont affichées, mais vous pouvez utiliser ce type de données dans le cadre d’un complément d’approbation de document. Si un utilisateur sélectionne un nom dans cette liste, votre code définit la valeur de la colonne du réviseur dans une liste de suivi de document. Un flux de travail associé à cette liste peut envoyer une notification de révision à cet utilisateur. Vous pouvez également enregistrer le nom sélectionné dans les paramètres du document. Dans ce cas, lorsqu’un utilisateur ouvre le document, vous pouvez afficher les contrôles dans le complément du volet Office uniquement si l’utilisateur actuel et l’utilisateur stocké dans les paramètres du document sont identiques. Pour plus d’informations, voir les sections suivantes :


- [Effectuer des opérations de base à l’aide de terminaux REST SharePoint 2013](http://msdn.microsoft.com/library/e3000415-50a0-426e-b304-b7de18f2f7d9%28Office.15%29.aspx)
    
- [Procédure : effectuer des opérations de base avec du code de bibliothèque JavaScript dans SharePoint 2013](http://msdn.microsoft.com/library/29089af8-dbc0-49b7-a1a0-9e311f49c826%28Office.15%29.aspx)
    
- [Conservation de l’état et des paramètres des compléments](../../docs/develop/persisting-add-in-state-and-settings.md)
    

1. Dans l’**Explorateur de solutions**, développez le dossier **MyTaskPaneAddin** et le dossier **Home**, puis sélectionnez le fichier **Home.html**.
    
    Le fichier Home.html s’ouvre dans l’éditeur de code.
    
2. Ajoutez le code HTML suivant sous le bouton `get-data-from-selection`.
    
```HTML
  <p>Select Reviewer:</p> <select class="select" id="select-reviewer" name="D1"> </select>
```

3. Sélectionnez le fichier **Home.js** pour ouvrir ce dernier dans l’éditeur de code.
    
4. Ajoutez les déclarations suivantes en haut du fichier Home.js.
    
```js
  var appWebURL; var web;
```

5. Remplacez la fonction  `Initialize` par le code suivant.
    
    Ce code effectue les tâches suivantes :
    
      - Charge les fichiers SP.Runtime.js et SP.js à l’aide de la fonction  `getScript` dans jQuery. Après le chargement des fichiers, votre programme a accès au modèle objet JavaScript pour SharePoint.
    
  - Charge l’objet de site Web actuel.
    
  - Appelle une fonction qui obtient tous les utilisateurs du site. Vous ajouterez le code de cette fonction à l’étape suivante.
    



```js
   // The initialize function must be run each time a new page is loaded Office.initialize = function (reason) { $(document).ready(function () { app.initialize(); var scriptbase = "/_layouts/15/"; $.getScript(scriptbase + "SP.Runtime.js", function () { $.getScript(scriptbase + "SP.js", function () { getAppWeb(function () { getSPUsers(populateUsersDropDown); }); }); }); function getAppWeb(functionToExecuteOnReady) { var context = SP.ClientContext.get_current(); web = context.get_web(); context.load(web); context.executeQueryAsync(onSuccess, onFailure); function onSuccess() { appWebURL = web.get_url(); functionToExecuteOnReady(); } function onFailure(sender, args) { app.initialize(); app.showNotification("Failed to connect to SharePoint. Error: " + args.get_message()); } } $('#get-data-from-selection').click(getDataFromSelection); }); };
```

6. Ajoutez le code suivant au bas du fichier Home.js.
    
    Ce code obtient la liste des utilisateurs du site Web à l’aide de l’interface REST fournie par SharePoint 2013. Ensuite, il remplit une liste déroulante avec les noms et les ID de chaque utilisateur.
    


```js
  function getSPUsers(functionToExecuteOnReady) { var url = appWebURL + "/../_api/web/siteUsers"; jQuery.ajax({ url: url, type: "GET", headers: { "ACCEPT": "application/json;odata=verbose" }, success: onSuccess, error: onFailure }); function onSuccess(data) { var results = data.d.results; functionToExecuteOnReady(results); } function onFailure(jaXHR, textStatus, errorThrown) { var error = textStatus + " " + errorThrown; app.showNotification(error); } } function populateUsersDropDown(results) { for (var i = 0; i < results.length; i++) { var IDTemp = results[i].Id; $('#select-reviewer').append("<option value='" + IDTemp + "'>" + results[i].Title + "</option>"); } }
```

7. Dans l’**Explorateur de solutions**, ouvrez le menu contextuel du fichier **AppManifest.xml** et sélectionnez **Concepteur de vues**.
    
8. Dans le concepteur, sélectionnez la page **Autorisations**.
    
9. Dans la liste déroulante située sous la colonne **Étendue**, sélectionnez l’élément **Web**.
    
10. Dans la liste déroulante située sous la colonne **Autorisation**, sélectionnez l’élément **Lecture**.
    

## Déboguer le complément du volet Office


Vous pouvez déboguer votre complément du volet Office en lançant le document ou en démarrant le SharePoint Add-in, puis en ouvrant un document à partir de la bibliothèque de documents.


### Débogage de votre complément du volet Office en lançant le document




 >**Remarque** : cette procédure ouvrant Excel, elle fonctionne uniquement lorsque Office est installé sur le système. Dans le cas contraire, un message d’erreur indique que « l’application associée à ce type de projet n’est pas installée sur cet ordinateur ».


1. Ouvrez le fichier Home.js dans l’éditeur de code, puis définissez un point d’arrêt en regard de la méthode `getDataFromSelection`.
    
2. Dans l’**Explorateur de solutions**, choisissez le nœud de projet **OfficeEnabledApp**.
    
3. Dans le menu **Affichage**, choisissez **Fenêtre Propriétés**.
    
4. Dans la fenêtre Propriétés, dans la liste déroulante  **Action de démarrage**, sélectionnez l’élément **Client de bureau Office**. Lorsque vous effectuez cette opération, une nouvelle propriété apparaît, **Démarrer le document**.
    
5. À partir de la liste déroulante **Démarrer le document**, choisissez l’élément **OfficeDocuments\TaskPaneApp.xlsx**.
    
6. Dans le menu **Déboguer**, sélectionnez **Démarrer le débogage**.
    
    Ce paramètre permet au classeur du complément du volet Office de s’afficher lorsque le complément s’exécute. Le classeur s’ouvre et le complément du volet Office s’affiche.
    
7. Dans le complément du volet Office, sélectionnez la liste déroulante **Sélectionner le réviseur** pour visualiser une liste des utilisateurs SharePoint.
    
8. Dans le classeur Excel, sélectionnez une cellule.
    
9. Dans le complément du volet Office, choisissez le bouton **Obtenir les données de la sélection**.
    
    L’exécution s’arrête au point d’arrêt que vous avez défini en regard de la méthode `getDataFromSelection`.
    

### Débogage de votre complément du volet Office en démarrant SharePoint




 >**Remarque** : cette procédure ouvre Excel Online. Cela fonctionne uniquement lorsque vous avez un compte Office 365. Voir la [procédure : Configurer un environnement de développement pour les compléments pour SharePoint dans Office 365](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29).


1. Ouvrez le fichier Home.js dans l’éditeur de code, puis définissez un point d’arrêt en regard de la méthode `getDataFromSelection`.
    
2. Dans l’**Explorateur de solutions**, choisissez le nœud de projet **OfficeEnabledApp**.
    
3. Dans le menu **Affichage**, choisissez **Fenêtre Propriétés**.
    
4. Dans la fenêtre Propriétés, dans la liste déroulante **Action de démarrage**, sélectionnez l’élément **Internet Explorer**.
    
5. Dans le menu **Déboguer**, sélectionnez **Démarrer le débogage**.
    
    Visual Studio ouvre SharePoint et affiche la bibliothèque **MyDocumentLibrary**.
    
6. Dans SharePoint, sous l’onglet **Fichiers**, choisissez **Nouveau document**. 
    
7. Accédez au classeur dans votre projet, MyTaskPaneApp.xlsx.
    
    Le classeur s’ouvre et le complément du volet Office s’affiche.
    
8. Vérifiez que le débogage de script est activé dans votre navigateur. Dans Internet Explorer, pour activer le débogage de script, ouvrez la boîte de dialogue **Options Internet**, choisissez l’onglet **Avancé**, puis désactivez les options **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.
    
9. Dans Visual Studio, dans le menu **Déboguer**, sélectionnez **Attacher au processus**.
    
10. Dans la boîte de dialogue **Attacher au processus**, choisissez tous les processus **Iexplore.exe** disponibles, puis sélectionnez le bouton **Attacher**.
    
11. Dans le complément du volet Office, sélectionnez la liste déroulante **Sélectionner le réviseur** pour visualiser une liste des utilisateurs SharePoint.
    
    Les données de la liste sont récupérées sur SharePoint à l’aide d’un appel REST.
    
12. Dans le classeur Excel, choisissez une cellule.
    
13. Dans le complément du volet Office, choisissez le bouton **Obtenir les données de la sélection**.
    
    L’exécution s’arrête au point d’arrêt que vous avez défini en regard de la méthode `getDataFromSelection`.
    
     >**Remarque** : si le classeur ne contient pas de données, vous pouvez en ajouter en choisissant **MODIFIER LE CLASSEUR**, **Modifier dans Excel Online** dans la barre d’outils du classeur.

## Empaqueter et publier le complément


Quand vous êtes prêt à empaqueter votre complément pour la publication, ouvrez l’Assistant **Publication des compléments SharePoint et Office**.


- Dans l’**Explorateur de solutions**, ouvrez le menu contextuel du projet de complément pour SharePoint, puis cliquez sur **Publier**.
    
    L’Assistant **Publication des compléments SharePoint et Office** apparaît. Pour plus d’informations, consultez [Publier des compléments pour SharePoint à l’aide de Visual Studio](http://msdn.microsoft.com/library/8137d0fa-52e2-4771-8639-60af80f693bb%28Office.15%29.aspx).
    

## Ressources supplémentaires


- [Instructions de conception pour les compléments Office](../../docs/design/add-in-design.md)
    
- [Cycle de vie du développement des compléments Office](../../docs/design/add-in-development-lifecycle.md)
    
- [Publier votre complément Office](../publish/publish.md)
    
- [Présentation de l’API JavaScript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Manifeste XML des compléments Office](../../docs/overview/add-in-manifests.md)
    
- [API et schémas de référence pour les compléments Office](../../reference/reference.md)
    
