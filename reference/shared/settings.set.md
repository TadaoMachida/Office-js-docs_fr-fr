

# Méthode Settings.set
Définit ou crée le paramètre spécifié.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans **|1.1|

```js
Office.context.document.settings.set(name, value);
```


## Paramètres



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **string**

&nbsp;&nbsp;&nbsp;&nbsp;Nom qui respecte la casse du paramètre à définir ou créer.

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **string**, **number**, **boolean**, **null**, **object** ou **array**

&nbsp;&nbsp;&nbsp;&nbsp;Spécifie la valeur à stocker.
    

## Remarques

La méthode **set** crée un paramètre ayant le nom spécifié s’il n’existe pas déjà. Sinon, elle définit un paramètre existant ayant le nom spécifié dans la copie en mémoire du conteneur des propriétés des paramètres. Une fois que vous avez appelé la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md), la valeur est stockée dans le document sous forme de représentation JSON sérialisée de son type de données. 2 Mo de stockage au maximum sont disponibles pour les paramètres de chaque complément.


 >**Important** : gardez à l’esprit que la méthode **Settings.set** concerne uniquement la copie en mémoire du conteneur des propriétés des paramètres. Pour vous assurer que les ajouts ou modifications apportés aux paramètres seront disponibles pour votre complément lors de la prochaine ouverture du document, après l’appel à la méthode **Settings.set** et avant la fermeture du complément, vous devez appeler la méthode **Settings.saveAsync** pour faire persister les paramètres du document.


## Exemple




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
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
|**Word**|v|v|v|

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
|1.1|Prise en charge supplémentaire des paramètres personnalisés dans les compléments du contenu Access.|
|1.0|Introduit|
