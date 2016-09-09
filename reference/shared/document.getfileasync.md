
# Méthode Document.getFileAsync
Renvoie l’intégralité du fichier de document sous forme de sections pouvant aller jusqu’à 4 194 304 octets (4 Mo). Pour des compléments pour iOS, la section de fichier est prise en charge jusqu'à 65 536 (64 Ko). Remarque : la spécification de la taille de section de fichier au-dessus de la limite autorisée entraîne une erreur interne. 

|||
|:-----|:-----|
|**Hôtes :**|Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Dernière modification dans le fichier**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|Spécifie le format dans lequel le fichier est renvoyé. Obligatoire.<br/><table><tr><th>Hôte</th><th>Type de fichier pris en charge</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>PowerPoint sur le bureau Windows</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr><tr><td>Word sur le bureau de Windows, MAC et iPad</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr></table>|**Modifié dans** 1.1, voir [Historique de prise en charge](#historique-de-prise-en-charge)|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _sliceSize_|**number**|Spécifie la taille de section souhaitée (en octets) pouvant aller jusqu’à 4 194 304 octets (4 Mo). Si aucune valeur n’est spécifiée, une taille de section par défaut de 4 194 304 octets (4 Mo) est utilisée. ||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getFileAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [File](../../reference/shared/file.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Pour les compléments exécutés dans des applications hôtes Office autres qu’Office pour iOS, la méthode **getFileAsync** prend en charge l’obtention de fichiers sous forme de sections pouvant aller jusqu’à 4 194 304 octets (4 Mo). Pour les compléments exécutés dans Office d’applications iOS, la méthode **getFileAsync** prend en charge l’obtention de fichiers sous forme de sections pouvant aller jusqu’à 65 536 octets (64 Ko).

Le paramètre _fileType_ peut être spécifié à l’aide des énumérations ou des valeurs de texte suivantes.


**FileType, énumération**


|**Énumération**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Renvoie l’intégralité du document (.docx, .pptx ou .xslx) au format Office Open XML (OOXML) sous forme de tableau d’octets.|
|Office.FileType.Pdf|"pdf"|Retourne l’intégralité du document au format PDF sous la forme d’un tableau d’octets.|
|Office.FileType.Text|"text"|Renvoie uniquement le texte du document sous forme de **chaîne**. |
Au maximum deux documents sont autorisés à se trouver en mémoire ; autrement, l’opération **getFileAsync** échoue. Utilisez la méthode [File.closeAsync](../../reference/shared/file.closeasync.md) pour fermer le fichier lorsque vous avez terminé de l’utiliser.


## Exemple : obtenir un document au format (« compressé ») Office Open XML

L’exemple suivant permet d’obtenir le document au format Office Open XML (« compressé ») sous forme de sections de 65 536 octets (64 Ko). Remarque : l’implémentation d’`app.showNotification` dans cet exemple provient du modèle Visual Studio pour les compléments Office.


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## Exemple : obtenir un document au format PDF

L’exemple suivant obtient le document au format PDF.


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||v||
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Fichier|
|**Niveau d’autorisation minimal**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1| Dans PowerPoint Online, **Office.FileType.Pdf** est désormais pris en charge en tant que paramètre _fileType_.|
|1.1| Dans PowerPoint Online, **Office.FileType.Compressed** est désormais pris en charge en tant que paramètre _fileType_.|
|1.1| Dans Word Online, **Office.FileType.Text** est désormais pris en charge en tant que paramètre _fileType_.|
|1.1| Dans Excel Online, **Office.FileType.Compressed** est désormais pris en charge en tant que paramètre _fileType_.|
|1.1| Dans Word Online, **Office.FileType.Compressed** et **Office.FileType.Pdf** sont désormais pris en charge pour le paramètre _fileType_.|
|1.1|Dans PowerPoint et Word sur Office pour iPad, toutes les valeurs **FileType** sont désormais prises en charge pour le paramètre _fileType_.|
|1.1|Dans Word et PowerPoint sur le bureau Windows, **Office.FileType.Pdf** est désormais pris en charge en tant que paramètre _fileType_.|
|1.0|Introduit|
