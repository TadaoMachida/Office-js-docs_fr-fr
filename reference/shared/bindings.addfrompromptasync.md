
# Méthode Bindings.addFromPromptAsync
 Affiche l’interface utilisateur qui permet à l’utilisateur de spécifier une sélection à lier.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Pas dans un ensemble|
|**Dernière modification**|1.1|

```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Spécifie le type de l’objet de liaison à créer. Obligatoire. Renvoie **null** si le type spécifié ne peut pas être forcé sur l’objet sélectionné.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _id_|**string**|Spécifie le nom unique à utiliser pour identifier le nouvel objet de liaison. Si aucun argument n’est transmis pour le paramètre _id_, le [Binding.id](../../reference/shared/binding.id.md) est généré automatiquement.||
| _promptText_|**string**|Spécifie la chaîne à afficher dans l’interface utilisateur d’invite qui indique à l’utilisateur quoi sélectionner. Limité à 200 caractères. Si aucun argument _promptText_ n’est transmis, un message invitant l’utilisateur à effectuer une sélection s’affiche.||
| _sampleData_|[TableData](../../reference/shared/tabledata.md)|Spécifie un tableau d’exemples de données affiché dans l’interface utilisateur d’invite comme exemple des types de champs (colonnes) qui peuvent être liés par votre complément. Les en-têtes indiqués dans l’objet **TableData** spécifient les étiquettes utilisées dans l’interface utilisateur de sélection de champs. Facultatif. **Remarque** : ce paramètre est utilisé uniquement dans les compléments pour Access. Il est ignorée si indiqué lors de l’appel de la méthode dans un complément pour Excel.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **addFromPromptAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Binding](../../reference/shared/binding.md) représentant la sélection spécifiée par l’utilisateur.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Ajoute un objet de liaison du type spécifié à la collection [Bindings](../../reference/shared/bindings.bindings.md), qui est identifiée à l’aide du paramètre _id_ indiqué. La méthode échoue si la sélection spécifiée est introuvable.


## Exemple




```js
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
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


|**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Pas dans un ensemble|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel dans Office pour iPad.|
|1.1|Dans les compléments pour Excel, vous pouvez créer une liaison de tableau (en transmettant _Office.BindingType.Table_ pour **bindingType**) pour une plage de cellules qui contient des données tabulaires même lorsque les données n’ont pas été ajoutées à la feuille de calcul sous forme de tableau dans l’interface utilisateur Excel (à l’aide des commandes **Insérer**  >  **Tableaux**  > **Tableau** ou **Accueil**  >  **Styles**  >  **Mettre sous forme de tableau**).|
|1.1|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access. |
|1.1|Prise en charge supplémentaire de la liaison à des données de matrice en tant que liaison de tableau dans les compléments pour Excel.|
|1.0|Introduit|
