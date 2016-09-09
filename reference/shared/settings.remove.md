

# Méthode Settings.remove
Supprime le paramètre spécifié.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans **|1.1|

```js
Office.context.document.settings.remove(name);
```


## Paramètres



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **string**

&nbsp;&nbsp;&nbsp;&nbsp;Nom respectant l’emploi des majuscules et minuscules pour le paramètre à supprimer.

    



## Remarques

 **null** est une valeur valide pour un paramètre. Ainsi, l’affectation de la valeur **null** au paramètre n’entraînera pas sa suppression du conteneur des propriétés des paramètres.


 >**Important** : gardez à l’esprit que la méthode **Settings.remove** concerne uniquement la copie en mémoire du conteneur des propriétés des paramètres. Pour faire persister la suppression du paramètre spécifié dans le document, après l’appel de la méthode **Settings.remove** et avant la fermeture du complément, vous devez appeler la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


## Exemple




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Paramètres|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de la création de paramètres personnalisés dans les compléments de contenu pour Access.|
|1.0|Introduit|
