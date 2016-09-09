
# Présentation du jeton d’identité Exchange
Découvrez le contenu d’un jeton d’identité d’Exchange 2013.



Le jeton d’identité d’authentification que le serveur Exchange envoie à votre complément Outlook est transparent pour votre complément ; vous n’avez pas à connaître son contenu pour l’envoyer à votre serveur. Mais lorsque vous écrivez le code du service web qui interagit avec votre complément Outlook, vous devez savoir ce que ce jeton contient.

## Qu’entend-on par jeton d’identité ?


Un jeton d’identité est une chaîne à codage URL base-64 autosignée par le serveur Exchange qui l’envoie. Le jeton n’est pas chiffré, et la clé publique que vous utilisez pour valider la signature est stockée sur le serveur Exchange qui a émis le jeton. Le jeton est composé de trois parties : un en-tête, une charge utile et une signature. Dans la chaîne du jeton, les différentes parties sont séparées par un caractère « . » pour simplifier le fractionnement du jeton.

Exchange 2013 utilise un jeton JWT (JSON Web Token) pour le jeton d’identité. Pour plus d’informations sur les jetons JWT, voir le [document préliminaire Internet sur JWT (JSON Web Token)](http://self-issued.info/docs/draft-goland-json-web-token-00.html).


### En-tête du jeton d’identité

L’en-tête identifie le jeton et permet au service web de reconnaître le type de jeton présenté. L’exemple suivant montre comment se présente l’en-tête du jeton.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }
```

Le tableau suivant décrit les parties de l’en-tête du jeton d’identité.


**Parties de l’en-tête du jeton d’identité**


|**Revendication**|**Valeur**|**Description**|
|:-----|:-----|:-----|
|typ|« JWT »|Identifie le jeton comme un jeton web JSON. Tous les jetons d’identité fournis par le serveur Exchange sont des jetons JWT.|
|alg|« RS256 »|L’algorithme de hachage qui est utilisé pour créer la signature. Tous les jetons fournis par le serveur Exchange utilisent algorithme RS-256.|
|x5t|Empreinte de certificat|L’empreinte X.509 du jeton.|

### Charge utile du jeton d’identité

La charge utile contient les revendications d’authentification qui identifient le compte de messagerie et identifient le serveur Exchange qui a envoyé le jeton. L’exemple suivant montre à quoi ressemble la section de charge utile.
```js

{ 
   "aud" : "https://mailhost.contoso.com/IdentityTest.html", 
   "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
   "nbf" : "1331579055", 
   "exp" : "1331607855", 
   "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
   "isbrowserhostedapp":"true",
"appctx" : { 
     "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" :         "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" 
     } 
}
```
Le tableau suivant répertorie les différentes parties de la charge utile du jeton d’identité.


**Parties de la charge utile du jeton d’identité**


|**Revendication**|**Description**|
|:-----|:-----|
|aud|L’URL du complément qui a demandé le jeton. Un jeton est uniquement valide s’il est envoyé à partir du complément qui s’exécute dans le navigateur du client. Si le complément utilise la version 1.1 du schéma de manifestes des Compléments Office, cette URL correspond à celle indiquée dans le premier élément  **SourceLocation**, sous le type de formulaire  **ItemRead** ou **ItemEdit**, selon celui qui apparaît en premier dans l’élément [FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) du manifeste de complément.|
|iss|Un identificateur unique du serveur Exchange qui a émis le jeton. Tous les jetons émis par ce serveur Exchange auront le même identificateur.|
|nbf|La date et l’heure de début de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970. |
|exp|La date et l’heure de fin de validité du jeton. La valeur correspond au nombre de secondes depuis le 1er janvier 1970.|
|appctxsender|Identificateur unique du serveur Exchange qui a envoyé le contexte de l’application.|
|isbrowserhostedapp|Indique si le complément est hébergé dans un navigateur.|
|appctx|Le contexte d’application du jeton. |
L’information dans la revendication appctx vous fournit l’adresse du compte de messagerie, ainsi qu’un identificateur unique pour le compte. Le tableau suivant répertorie les différentes parties de la revendication appctx.



|**Partie de la revendication appctx**|**Description**|
|:-----|:-----|
|msexchuid|Un identificateur unique associé au compte de messagerie et au serveur Exchange.|
|version|Numéro de version du jeton. Pour tous les jetons fournis par un serveur qui exécute Exchange 2013, la valeur est « ExIdTok.V1 ».|
|amurl|URL du document de métadonnées d’authentification qui contient la clé publique du certificat X.509 qui est utilisée pour signer le jeton. Pour plus d’informations sur l’utilisation du document de métadonnées d’authentification, voir [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md).|

### Signature du jeton d’identité

La signature est créée par hachage des sections d’en-tête et de charge utile avec l’algorithme spécifié dans l’en-tête et en utilisant le certificat X509 autosigné situé sur le serveur à l’emplacement spécifié dans la charge utile. Votre service web peut valider cette signature pour contribuer à assurer que le jeton d’identité provient bien du serveur prévu pour son envoie.


## Ressources supplémentaires



- [Authentifier un complément Outlook à l’aide de jetons d’identité Exchange](../outlook/authentication.md)
    
- [Appeler un service à partir d’un complément Outlook à l’aide d’un jeton d’identité dans Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Utiliser la bibliothèque de validation des jetons Exchange](../outlook/use-the-token-validation-library.md)
    
- [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md)
    
