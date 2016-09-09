
# GoToType, énumération
Spécifie le type d’emplacement ou d’objet auquel accéder.

|||
|:-----|:-----|
|**Hôtes :**|Excel, PowerPoint, Word|
|**Ajouté dans**|1.1|

```js
Office.GoToType
```


## Membres


**Valeurs**


|**Énumération**|**Valeur**|**Description**|**Clients pris en charge**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|« binding »|Accède à un objet de liaison en utilisant l’ID de liaison spécifié.|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|Accède à un élément à l’aide du nom de cet élément, tel que le nom affecté à un tableau ou à une plage. Dans Excel, vous pouvez utiliser n’importe quelle référence structurée pour une plage ou un tableau nommé(e) : "Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|« slide »|Accède à une diapositive en utilisant l’ID spécifié.|PowerPoint|
|Office.GoToType.Index|« index »|Accès à l’index spécifié par numéro de diapositive ou énumération :</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.


Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Introduit|
