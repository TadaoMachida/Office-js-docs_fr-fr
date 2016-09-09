
# BindingType, énumération
 Spécifie le type de l’objet de liaison qui doit être retourné.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification**|1.1|

```
Office.BindingType
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|Données tabulaires sans ligne d’en-tête. Les données sont renvoyées sous forme d’un tableau de tableaux, par exemple : ` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|Données tabulaires avec une ligne d’en-tête. Les données sont renvoyées en tant qu’objet [TableData](../../reference/shared/tabledata.md).|
|Office.BindingType.Text|"text"|Texte brut. Les données sont retournées sous forme de suite de caractères.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|v|||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de la liaison de données de tableau dans les compléments pour Access.|
|1.0|Introduites|
