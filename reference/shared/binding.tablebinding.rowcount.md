
# Propriété tableBinding.rowCount
Obtient le nombre de lignes du tableau, sous forme de valeur entière.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Dernière modification dans la sélection**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## Valeur renvoyée

Nombre de lignes de l’objet [TableBinding](../../reference/shared/binding.tablebinding.md) spécifié.


## Remarques

Lorsque vous insérez un tableau vide en sélectionnant une seule ligne dans Excel 2013 et Excel Online (à l’aide de l’option **Tableau** sous l’onglet **Insertion**), les applications hôtes Office créent une ligne unique d’en-têtes suivie par une seule ligne vide. Cependant, si le script de votre complément crée une liaison pour ce nouveau tableau inséré (par exemple, à l’aide de la méthode [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)), puis vérifie la valeur de la propriété **rowCount**, la valeur renvoyée variera en fonction de l’ouverture de la feuille de calcul dans Excel 2013 ou dans Excel Online.


- Dans Excel, **rowCount** renvoie 0 (la ligne vide qui suit les en-têtes n’est pas comptabilisée).
    
- Dans Excel Online, **rowCount** renvoie 1 (la ligne vide qui suit les en-têtes est comptabilisée).
    
Vous pouvez contourner cette différence dans votre script en vérifiant si `rowCount == 1` et, si tel est le cas, en vérifiant si la ligne contient toutes les chaînes vides.

Dans les compléments de contenu pour Access, pour des raisons de performances, la propriété **rowCount** renvoie toujours -1.


## Exemple




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|TableBindings|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
