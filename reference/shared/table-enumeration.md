
# Table, énumération
Spécifie les valeurs énumérées de la propriété `cells:` dans le paramètre _cellFormat_ des [méthodes de mise en forme de tableau](../../docs/excel/format-tables-in-add-ins-for-excel.md).

|||
|:-----|:-----|
|**Hôtes :**|Excel|
|**Ajouté**|1.1|

```
Office.Table
```

## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.Table.All|« tout »|Le tableau entier, y compris les totaux, les données et les en-têtes de colonne (le cas échéant).|
|Office.Table.Data|« données »|Uniquement les données (sans les en-têtes ni les totaux).|
|Office.Table.Headers|« en-têtes »|Uniquement la ligne d’en-tête|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique les énumérations prises en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel dans Office pour iPad.|
|1.1|Introduit|
