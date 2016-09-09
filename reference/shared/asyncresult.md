
# AsyncResult, objet
Objet qui encapsule le résultat d’une requête asynchrone, y compris les informations d’état et d’erreur quand la demande a échoué.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```
AsyncResult
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|Obtient l’élément défini par l’utilisateur transmis au paramètre facultatif _asyncContext_ de la méthode appelée dans le même état que celui dans lequel il a été transmis.|
|**[erreur](../../reference/shared/asyncresult.error.md)**|Obtient un objet **Error** qui fournit une description de l’erreur, si une erreur s’est produite.|
|**[statut](../../reference/shared/asyncresult.status.md)**|Obtient l’état de l’opération asynchrone.|
|**[value](../../reference/shared/asyncresult.value.md)**|Obtient la charge utile ou le contenu de l’opération asynchrone, le cas échéant.|

## Remarques

Quand la fonction que vous transmettez au paramètre _callback_ pour une méthode « Async » s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

L’exemple suivant est applicable aux compléments de contenu et de volet de tâches. Il illustre un appel à la méthode [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) de l’objet **Document**.




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

La fonction anonyme transmise comme argument _callback_ (`function (result){...}`) a un seul paramètre nommé _result_ qui donne accès à un objet **AsyncResult** quand la fonction s’exécute. Quand l’appel à la méthode **getSelectedDataAsync** est terminé, la fonction de rappel s’exécute et la ligne de code suivante accède à la propriété **value** de l’objet **AsyncResult** pour renvoyer les données sélectionnées dans le document :

 `var dataValue = result.value;`

Notez que d’autres lignes de code de la fonction utilisent le paramètre _result_ de la fonction de rappel pour accéder aux propriétés **status** et **error** de l’objet **AsyncResult**.

L’objet **AsyncResult** est disponible à partir de la fonction transmise comme argument au paramètre _callback_ des méthodes suivantes :



|**Objet parent**|**Méthode**|
|:-----|:-----|
|**Document** (Excel, PowerPoint, Project et Word uniquement)|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings** (Excel et Word uniquement)|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding** (Excel et Word uniquement)|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding** (Excel et Word uniquement)||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings** (Excel, PowerPoint et Word uniquement)|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode** (Word uniquement)|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart** (Word uniquement)|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts** (Word uniquement)|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings** (Word uniquement)|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox** (Outlook uniquement)|[getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||[makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (Outlook uniquement)|[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (Outlook uniquement)|[loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (Outlook uniquement)|[saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).



| |**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Projet**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
