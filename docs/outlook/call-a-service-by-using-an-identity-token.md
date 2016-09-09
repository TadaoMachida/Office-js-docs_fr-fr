
# Appeler un service à partir d’un complément Outlook à l’aide d’un jeton d’identité dans Exchange

Un jeton d’identité fournit un identificateur unique pour chacun de vos clients, que vous pouvez utiliser afin de personnaliser le service que vous fournissez. Votre code peut demander au serveur Exchange un jeton d’identité en utilisant une méthode asynchrone qui renvoie une chaîne à votre complément Outlook. La chaîne contient un jeton d’identité JSON Web Token (JWT). Votre complément n’a pas besoin de décompresser le jeton. Au lieu de cela, il le transmet à votre service web de façon que ce dernier puisse authentifier la demande du complément.

Le service web qui prend en charge votre complément doit s’exécuter sur le serveur qui héberge les fichiers sources HTML et JavaScript du complément. Ceci évite les erreurs de script intersite. Votre serveur peut transmettre la demande à d’autres services web si votre application l’exige.

Il est simple d’ajouter un jeton d’identité à la demande de service que votre complément envoie : demandez le jeton, utilisez-le, puis utilisez la réponse du service web. Voici ce à quoi cela ressemble dans le cas d’un document XML simple que vous envoyez à votre serveur avec la méthode **XmlHttpRequest**.

## Demande d’un jeton à votre serveur Exchange


Cette méthode d’initialisation simple d’un complément utilise la méthode  **getUserIdentityTokenAsync** pour demander un jeton d’identité au serveur Exchange. Le paramètre _getUserIdentityToken_ est la fonction qui est appelée au retour de la demande asynchrone adressée au serveur. Voir l’étape suivante pour la méthode de rappel.


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## Utilisation du jeton d’identité


La fonction de rappel pour la méthode  **getUserIdentityTokenAsync** possède un paramètre qui contient le jeton d’identité de l’utilisateur dans sa propriété **value**.

Cette fonction de rappel crée un objet  **XMLHttpRequest** pour appeler le service web. Affectez à la propriété **onreadystatechange** de l’objet **XMLHttpRequest** le nom de la fonction qui doit s’exécuter lorsque votre complément obtient une réponse du service web.




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## Utilisation de la réponse du service web


Il s’agit d’une autre fonction simple qui traite la réponse du service web. Elle respecte le modèle standard pour les fonctions de rappel  **XHMHttpResponse**. Elle attend que la réponse entière arrive du service web et place le contenu de la réponse dans l’interface utilisateur du complément. La réponse que cette fonction analyse est la réponse du service web. Pour plus d’informations sur cette réponse, voir [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md). 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## Exemple : appel d’un service Web avec des jetons d’identité


Les jetons d’identité fournissent des informations d’identité sur le client qui appelle votre service auprès d’un service web s’exécutant sur votre serveur. Pour utiliser les jetons d’identité, vous avez besoin de ce qui suit :


- Un complément Outlook qui demande un jeton d’identité au serveur Exchange et le transmet à votre service web. Les informations de cette rubrique vous aideront à créer ce complément.
    
- Un service web s’exécutant sur le serveur qui fournit l’interface utilisateur de votre complément qui valide le jeton d’identité. Vous trouverez les informations dont vous avez besoin pour créer le service web dans l’une des rubriques suivantes :
    
      - [Utiliser la bibliothèque de validation des jetons Exchange](../outlook/use-the-token-validation-library.md) -- Si vous utilisez la bibliothèque de validation que nous fournissons.
    
  - [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md) -- Si vous écrivez votre propre code de validation.
    

### Code pour l’exemple de complément


Les fichiers suivants sont nécessaires pour le complément décrit dans cet article :


- IdentityTest.js – Fichier JavaScript qui fournit la logique métier pour le complément.
    
- IdentityTest.html – Fichier HTML qui fournit l’interface utilisateur pour le complément.
    
Vous aurez également besoin du service web Identity Test. Pour plus d’informations sur ce service web, voir [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md).


#### IdentityTest.js

L’exemple suivant présente le fichier IdentityTest.js.


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### IdentityTest.html

L’exemple suivant présente le fichier IdentityTest.html.


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## Étapes suivantes


Une fois que vous savez comment demander un jeton d’identité, vous devez utiliser le jeton du côté serveur de la requête. Les articles suivants vous aideront à démarrer :


- [Utiliser la bibliothèque de validation des jetons Exchange](../outlook/use-the-token-validation-library.md)
    
- [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md)
    
- [Authentifier un utilisateur avec un jeton d’identité pour Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## Ressources supplémentaires



- [Authentifier un complément Outlook à l’aide de jetons d’identité Exchange](../outlook/authentication.md)
    
- [Présentation du jeton d’identité Exchange](../outlook/inside-the-identity-token.md)
    
