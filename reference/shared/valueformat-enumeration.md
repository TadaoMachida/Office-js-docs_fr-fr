
# ValueFormat, énumération
Spécifie si les valeurs (telles que les nombres et les dates) retournées par la méthode appelée sont retournées avec leur mise en forme appliquée.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Project, Word|
|**Ajouté dans**|1,0|

```
Office.ValueFormat
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|Retourne les données mises en forme.|
|Office.ValueFormat.Unformatted|"unformatted"|Retourne les données non mises en forme.|

## Remarques

Par exemple, si le paramètre _valueFormat_ est spécifié en tant que `"formatted"`, le format de devise d’un nombre ou le format de date (jj/mm/aa) est conservé dans l’application hôte. Si le paramètre _valueFormat_ est spécifié en tant que `"unformatted"`, la date est renvoyée dans sa forme numérique séquentielle sous-jacente.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Projet**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.0|Introduit|
