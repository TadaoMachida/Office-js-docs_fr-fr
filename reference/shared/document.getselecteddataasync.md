
# Méthode Document.getSelectedDataAsync
Lit les données contenues dans la sélection actuelle du document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Disponible dans les ensembles de ressources requis**|Selection|
|**Dernière modification dans la sélection**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>Prise en charge d’hôte</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (chaîne)</td><td>Excel, Excel Online, PowerPoint, PowerPoint Online, Word et Word Online uniquement</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (tableau de tableaux)</td><td>Excel, Word et Word Online uniquement</td></tr><tr><td><b>Office.CoercionType.Table</b> (objet [TableData](../../reference/shared/tabledata.md))</td><td>Access, Excel, Word et Word Online uniquement</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>Données au format HTML</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>Word et Word Online uniquement</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>PowerPoint et PowerPoint Online uniquement</td></tr></table>|Type de structure de données à retourner. Requis.||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>Spécifie si le résultat doit être renvoyé avec ses valeurs numériques ou de date mises en forme ou non.</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>Spécifie si le filtrage doit être appliqué lorsque les données sont récupérées. Facultatif.</td><td>Ce paramètre est ignoré dans les documents Word.</td></tr><tr><td><i>asyncContext</i></td><td><b>tableau</b>, <b>booléen</b>, <b>null</b>, <b>numérique</b>, <b>objet</b>, <b>chaîne</b> ou <b>non défini</b></td><td>Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet <b>AsyncResult</b> sans être modifié.</td><td></td></tr></table>|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getSelectedDataAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder aux valeurs de la sélection actuelle, lesquelles sont renvoyées dans la structure de données ou au format que vous avez spécifié(e) avec le paramètre _coercionType_. (Voir **Remarques** pour plus d’informations sur le forçage de type de données.)|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Dans votre volet de tâches ou votre complément de contenu, utilisez la méthode **getSelectedDataAsync** pour écrire le script qui lit les données sélectionnées par l’utilisateur dans un document, une feuille de calcul, une présentation ou un projet. Par exemple, une fois qu’un utilisateur a sélectionné du contenu dans un document Word, vous pouvez utiliser la méthode **getSelectedDataAsync** pour lire cette sélection, puis la soumettre à un service Web sous forme de requête ou de toute autre opération.

Après la lecture de la sélection, vous pouvez également utiliser les méthodes [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) et [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) de l’objet **Document** pour [mettre à jour la sélection ou ajouter un gestionnaire d’événements](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md), afin de détecter si l’utilisateur modifie la sélection.

La méthode **getSelectedDataAsync** peut seulement lire les éléments sélectionnés tant qu’ils sont actifs. Dans les compléments pour Word et Excel, si vous devez créer une association persistante de lecture et d’écriture vers la sélection de l’utilisateur, utilisez plutôt la méthode [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) pour [créer une liaison vers cette sélection](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).

Utilisez le paramètre _coercionType_ de la méthode **getSelectedDataAsync** pour spécifier la structure ou le format des données sélectionnées en cours de lecture.



|**Paramètre _coercionType_ spécifié**|**Données renvoyées**|**Données renvoyées**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** ou `"text"`|Chaîne.|Chaîne<br/><br/> **Remarque** : dans Excel, même si un sous-ensemble d’une cellule est sélectionné, l’intégralité du contenu de la cellule est renvoyé.|
|**Office.CoercionType.Matrix** ou `"matrix"`|Tableau de tableaux. Par exemple, ` [['a','b'], ['c','d']]` pour une sélection de deux lignes dans deux colonnes.|Objet TableData pour lire un tableau avec des en-têtes.|
|**Office.CoercionType.Table** ou `"table"`|Objet [TableData](../../reference/shared/tabledata.md) pour la lecture d’un tableau avec des en-têtes.|Objet TableData pour lire un tableau avec des en-têtes.|
|**Office.CoercionType.Html** ou `"html"`|Au format HTML.|Données au format HTML|
|**Office.CoercionType.Ooxml** ou `"ooxml"`|Au format Open Office XML (OpenXML).|Données au format HTML<br/><br/> **Conseil** : Lorsque vous développez le code de votre complément, vous pouvez utiliser le `"ooxml"`_coercionType_ de la méthode **getSelectedDataAsync** pour voir comment le contenu que vous sélectionnez dans un document Word est défini en tant que balises OpenXML. Ensuite, utilisez ces balises dans le paramètre de données de la méthode [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) pour écrire du contenu avec cette mise en forme ou structure dans un document. Par exemple, vous pouvez [insérer une image dans un document](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx) au format OpenXML.|
|**Office.CoercionType.SlideRange** ou "slideRange"|Objet JSON qui contient un tableau nommé « slides » qui contient les ID, les titres et les index des diapositives sélectionnées.  **Remarque :** Pour sélectionner plusieurs diapositives, l’utilisateur doit modifier la présentation dans l’affichage **Normal**, **Mode Plan** ou **Trieuse de diapositives**. En outre, cette méthode n’est pas prise en charge dans **Modes Masques**. Par exemple, `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` pour une sélection de deux diapositives.|PowerPoint uniquement.|
Si la structure de données de la sélection ne correspond pas au _coercionType_ spécifié, la méthode **getSelectedDataAsync** tentera de forcer le type de données sur ce type ou cette structure. Si le type **Office.CoercionType** spécifié ne peut pas être forcé sur la sélection, la propriété **AsyncResult.status** renvoie `"failed"`.


## Exemple

Si la structure de données de la sélection ne correspond pas au paramètre coercionType spécifié, la méthode getSelectedDataAsync tente de forcer le type de données dans ce type ou cette structure. Si ce forçage échoue pour le type Office.CoercionType que vous avez indiqué, la propriété AsyncResult.status renvoie "failed".


-  **Transmettre une fonction de rappel anonyme** qui lit la valeur de la sélection actuelle au paramètre _callback_ de la méthode **getSelectedDataAsync**.
    
-  **Lire la sélection** en tant que texte, sans mise en forme et non filtré.
    
-  **Afficher la valeur** sur la page du complément.
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
            }            
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Projet**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Selection|
|**Niveau d’autorisation minimal**|[Niveau d’autorisation minimal](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1| Dans Word Online, **Office.CoercionType.Matrix** et **Office.CoercionType.Table** sont désormais pris en charge pour le paramètre _coercionType_.|
|1.1|Dans Excel, PowerPoint et Word dans Office pour iPad, le même niveau de prise en charge que dans Excel, PowerPoint et Word est désormais disponible sur le bureau Windows.|
|1.1| Dans Word Online, **Office.CoercionType.Text** est désormais pris en charge en tant que paramètre _coercionType_.|
|1.1|Dans les compléments de contenu pour PowerPoint, vous pouvez obtenir les ID, les titres et les index de la plage de diapositives sélectionnée en transmettant **Office.CoercionType.SlideRange** en tant que paramètre _coercionType_ de la méthode **getSelectedDataAsync**. Voir la rubrique sur la méthode [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) pour obtenir un exemple d’utilisation de cette valeur afin de naviguer jusqu’à la diapositive actuellement sélectionnée.|
|1.0|Introduit|
