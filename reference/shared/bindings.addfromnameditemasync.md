
# Méthode Bindings.addFromNamedItemAsync
Ajoute une liaison à un élément nommé dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Dernière modification**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|Nom de l’élément nommé. Requis.||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Spécifie le type de l’objet de liaison à créer. Obligatoire. Renvoie **null** si le type spécifié ne peut pas être forcé sur l’objet sélectionné.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _id_|**string**|Spécifie le nom unique à utiliser pour identifier le nouvel objet de liaison. Si aucun argument n’est transmis pour le paramètre _id_, le [Binding.id](../../reference/shared/binding.id.md) est généré automatiquement.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **addFromNamedItemAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Binding](../../reference/shared/binding.md) qui représente l’élément nommé spécifié.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

 **Pour Excel**, le paramètre _itemName_ peut faire référence à une plage nommée ou à un tableau.

Par défaut, l’ajout d’un tableau dans Excel entraîne l’affectation du nom « Tableau1 » pour le premier tableau que vous ajoutez, « Tableau2 » pour le deuxième tableau que vous ajoutez, et ainsi de suite. Pour affecter un nom significatif à un tableau dans l’interface utilisateur d’Excel, servez-vous de la propriété **Nom du tableau** sous l’onglet **Outils de tableau ** Création| du ruban.


 >**Remarque**  Dans Excel, lors de la spécification d’un tableau comme élément nommé, vous devez entièrement qualifier le nom pour inclure le nom de la feuille de calcul dans le nom du tableau dans ce format :  `"Sheet1!Table1"`

 **Pour Word**, le paramètre _itemName_ fait référence à la propriété **Titre** d’un contrôle de contenu **Texte enrichi**. (Vous ne pouvez pas définir de liaison avec des contrôles de contenu autres que **Texte enrichi**.)

Par défaut, un contrôle de contenu n’a aucune valeur  **Titre** affectée. Pour affecter un nom significatif dans l’interface utilisateur de Word, après l’insertion d’un contrôle de contenu **Texte enrichi** à partir du groupe **Contrôles** sur l’onglet **Développeur** du ruban, utilisez la commande **Propriétés** du groupe **Contrôles** pour afficher la boîte de dialogue **Propriétés du contrôle de contenu**. Définissez ensuite la propriété  **Titre** du contrôle de contenu sur le nom auquel vous souhaitez faire référence à partir de votre code.


 >**Remarques**  Dans Word, s’il existe plusieurs contrôles de contenu **Texte enrichi** avec la même valeur de propriété **Titre** (le même nom) et que vous essayez de lier l’un de ces contrôles de contenu à cette méthode (en spécifiant son nom comme paramètre _itemName_), l’opération échoue.


## Exemple

L’exemple suivant ajoute une liaison à l’élément nommé `myRange` dans Excel sous forme de liaison « matrix » (matrice), puis affecte à l’[id](../../reference/shared/binding.id.md) de la liaison la valeur `myMatrix`.


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

L’exemple suivant ajoute une liaison à l’élément nommé `Table1` dans Excel sous forme de liaison « table », puis affecte à l’**id** de la liaison la valeur `myTable`.




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

L’exemple suivant crée une liaison de texte dans Word vers un contrôle de contenu de texte enrichi nommé  `"FirstName"`, attribue l’ **id**`"firstName"`, puis affiche cette information.




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
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

||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|MatrixBindings, TableBindings, TextBindings|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Dans les compléments pour Excel, vous pouvez créer une liaison de tableau (en transmettant _Office.BindingType.Table_ pour **bindingType**) pour une plage de cellules qui contient des données tabulaires même lorsque les données n’ont pas été ajoutées à la feuille de calcul sous forme de tableau (à l’aide des commandes **Insérer**  >  **Tableaux**  > **Tableau** ou **Accueil**  >  **Styles**  >  **Mettre sous forme de tableau**).|
|1.1|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access. |
|1.0|Introduit|

## Voir aussi



#### Autres ressources


[Lier des régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
