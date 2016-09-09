

# Méthode Office.select
Crée une promesse de retour d’une liaison en fonction de la chaîne de sélecteur passée.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans les [ensembles de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Dernière modification dans **|1.1|

```js
Office.select(str, onError);
```


## Paramètres


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **string**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Chaîne de sélecteur à analyser et pour laquelle une promesse doit être créée.

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**. Facultatif.
    

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _onError_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel. Si l’opération a échoué, utilisez la propriété [AsyncResult.error](../../reference/shared/asyncresult.error.md) pour accéder à un objet [Error](../../reference/shared/error.md) qui fournit des informations sur l’erreur.


## Remarques

La méthode **Office.select** permet d’accéder à une promesse d’objet [Binding](../../reference/shared/binding.md) qui tente de renvoyer la liaison spécifiée quand ses méthodes asynchrones sont appelées.

Formats pris en charge : « bindings# _bindingId_ », qui retourne un objet **Binding** pour la liaison ayant l’[ID](../../reference/shared/binding.id.md) `bindingId`. Pour plus d’informations, voir [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings) et [Lier des régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


 >**Remarque** : Si la promesse de la méthode **select** renvoie un objet **Binding**, cet objet expose uniquement les quatre méthodes suivantes de l’objet [Binding](../../reference/shared/binding.md) : [getDataAsync](../../reference/shared/binding.getdataasync.md), [setDataAsync](../../reference/shared/binding.setdataasync.md), [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) et [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md). Si la promesse ne peut pas renvoyer un objet **Binding**, le rappel _onError_ peut être utilisé pour accéder à un objet [asyncResult.error](../../reference/shared/asyncresult.error.md) dans le but d’obtenir plus d’informations. Si vous devez appeler un membre de l’objet **Binding** autre que les quatre méthodes exposées par la promesse de l’objet **Binding** renvoyé par la méthode **select**, utilisez plutôt la méthode [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) en employant la propriété [Document.bindings](../../reference/shared/document.bindings.md) et la méthode [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) pour récupérer l’objet **Binding**.


## Exemple

L’exemple de code suivant utilise la méthode **select** pour récupérer une liaison avec l’**id** « `cities` » à partir de la collection **Bindings**, puis appelle la méthode [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) afin d’ajouter un gestionnaire d’événements pour l’événement [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) de la liaison.


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
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
|**Disponible dans les ensembles de ressources requis**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Niveau d’autorisation minimal**|[ReadDocument (ReadAllDocument pour Open Office XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad|
|1.1|Ajout de l’utilisation de la méthode **select** pour renvoyer les liaisons de tableau créées dans les compléments de contenu pour Access.|
|1.0|Introduit|
