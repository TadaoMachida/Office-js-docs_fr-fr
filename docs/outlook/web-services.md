
# Appeler des services web à partir d’un complément Outlook

Votre complément peut utiliser les services web Exchange (EWS) d’un ordinateur exécutant Exchange Server 2013, un service web disponible sur le serveur qui fournit l’emplacement source de l’interface utilisateur du complément ou un service web disponible sur Internet. Cette rubrique fournit des exemples expliquant comment un complément Outlook peut demander des informations à partir d’EWS.

La méthode d’appel d’un service Web dépend de l’emplacement de ce dernier. Le tableau 1 répertorie les méthodes d’appel d’un service Web en fonction de l’emplacement.


**Tableau 1. Méthodes d’appel de services web à partir d’un complément Outlook**


|**Emplacement des services web**|**Méthode d’appel du service Web**|
|:-----|:-----|
|Serveur Exchange qui héberge la boîte aux lettres cliente|Utilisez la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) pour appeler les opérations EWS qui permettent d'ajouter des compléments de prise en charge. Le serveur Exchange qui héberge la boîte aux lettres expose également EWS.|
|Serveur web qui fournit l’emplacement source de l’interface utilisateur du complément|Appelez le service web au moyen des techniques JavaScript standard. Le code JavaScript présent dans le cadre de l’interface utilisateur s’exécute dans le contexte du serveur web qui fournit l’interface utilisateur. Il est donc capable d’appeler les services web sur ce serveur sans provoquer d’erreur de script intersite.|
|Tous les autres emplacements|Créez un proxy pour le service web sur le serveur web qui fournit l’emplacement source de l’interface utilisateur. Si vous n’indiquez pas de proxy, votre complément ne s’exécutera pas en raison d’erreurs de script intersites. L’un des moyens de fournir un proxy consiste à utiliser JSON/P. Pour plus d’informations, voir [Confidentialité et sécurité pour les compléments Office](../../docs/develop/privacy-and-security.md).|

## Utilisation de la méthode makeEwsRequestAsync pour accéder aux opérations EWS


Vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) pour effectuer une demande EWS auprès du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.

EWS prend en charge en charge différentes opérations sur un serveur Exchange, par exemple, les opérations au niveau de l’élément pour copier, rechercher, mettre à jour ou envoyer un élément, et les opérations au niveau du dossier pour créer, obtenir ou mettre à jour un dossier. Pour exécuter une opération EWS, créez une demande SOAP XML pour cette opération. Une fois l’opération terminée, vous obtenez une réponse SOAP XML qui contient les données correspondant à l’opération. Les demandes et les réponses SOAP EWS suivent le schéma défini dans le fichier Messages.xsd. Comme d’autres fichiers de schéma EWS, le fichier Message.xsd se trouve dans le répertoire virtuel IIS qui héberge EWS. 

Pour utiliser la méthode  **makeEwsRequestAsync** pour initier une opération EWS, indiquez les éléments suivants :


- Code XML pour la demande SOAP pour cette opération EWS, en tant qu’argument du paramètre  _data_
    
- Méthode de rappel (en tant qu’argument  _callback_)
    
- Données d’entrée facultatives pour cette méthode de rappel (en tant qu’argument  _userContext_)
    
Une fois la demande SOAP EWS terminée, Outlook appelle la méthode de rappel avec un argument, qui est un objet [AsyncResult](../../reference/outlook/simple-types.md). La méthode de rappel peut accéder à deux propriétés de l’objet  **AsyncResult** : la propriété  **value**, qui contient la réponse SOAP XML de l’opération EWS, et éventuellement la propriété  **asyncContext**, qui contient les données transmises en tant que paramètre  **userContext**. En règle générale, la méthode de rappel analyse ensuite le code XML dans la réponse SOAP pour obtenir les informations pertinentes et traite ces informations comme il se doit.


## Conseils pour l’analyse des réponses EWS


Lors de l’analyse d’une réponse SOAP à partir d’une opération EWS, notez les problèmes dépendants du navigateur suivants :


- Spécifiez le préfixe de nom de balise lorsque vous utilisez la méthode DOM  **getElementsByTagName**, pour inclure la prise en charge d’Internet Explorer.
    
     **getElementsByTagName** se comporte différemment selon le type de navigateur. Par exemple, une réponse EWS peut contenir le code XML suivant (mis en forme et abrégé à des fins d’affichage) :
    
```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
```

 Un code tel que le suivant fonctionnera dans un navigateur tel que Chrome pour obtenir le code XML entouré par les balises  **ExtendedProperty** :

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
```


   
 Sur Internet Explorer, vous devez inclure le préfixe `t:` du nom de balise, comme indiqué ci-dessous :

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
```

- Utilisez la propriété DOM  **textContent** pour obtenir le contenu d’une balise dans une réponse EWS, comme indiqué ci-dessous :
    
```
      content = $.parseJSON(value.textContent);
```

 D’autres propriétés telles que  **innerHTML** peuvent ne pas fonctionner sur Internet Explorer pour certaines balises dans les réponses EWS.
    

## Exemple


L’exemple suivant appelle  **makeEwsRequestAsync** pour utiliser l’opération [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) afin d’obtenir l’objet d’un élément. Cet exemple comprend les trois fonctions suivantes :


-  `getSubjectRequest` -- Prend un ID d’élément comme entrée et retourne le XML pour la demande SOAP qui appelle **GetItem** pour l’élément spécifié.
    
-  `sendRequest` -- Appelle `getSubjectRequest` pour obtenir la demande SOAP pour l’élément sélectionné, puis passe la demande SOAP et la méthode de rappel `callback` à **makeEwsRequestAsync** pour obtenir l’objet de l’élément spécifié.
    
-  `callback` -- Traite la réponse SOAP qui comprend l’objet et d’autres informations sur l’élément spécifié.
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## Opérations EWS prises en charge par les compléments


Les compléments Outlook peuvent accéder à un sous-ensemble d’opérations disponibles dans EWS par le biais de la méthode  **makeEwsRequestAsync**. Si vous ne connaissez pas les opérations EWS et ne savez pas comment utiliser la méthode  **makeEwsRequestAsync** pour accéder à une opération, commencez avec un exemple de demande SOAP pour personnaliser votre argument _data_. Voici des explications sur la manière d’utiliser la méthode  **makeEwsRequestAsync** :


1. Dans le XML, remplacez les ID d’éléments et les attributs d’opération EWS par les valeurs appropriées.
    
2. Intégrez la demande SOAP en tant qu’argument pour le paramètre  _data_ de **makeEwsRequestAsync**.
    
3. Spécifiez une méthode de rappel et appelez  **makeEwsRequestAsync**.
    
4. Dans la méthode de rappel, vérifiez les résultats de l’opération dans la réponse SOAP.
    
5. Utilisez les résultats de l’opération EWS en fonction de vos besoins.
    
Le tableau suivant répertorie les opérations EWS prises en charge par les compléments. Pour afficher des exemples de demandes et réponses SOAP, choisissez le lien correspondant à chaque opération. Pour plus d’informations sur les opérations EWS, voir [Opérations EWS dans Exchange](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx).


**Tableau 2. Opérations EWS prises en charge**


|**Opération EWS**|**Description**|
|:-----|:-----|
|[CopyItem Operation](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|Copie les éléments spécifiés et place les nouveaux éléments dans un dossier spécifique dans la banque d’informations Exchange.|
|[CreateFolder Operation](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|Crée les dossiers dans l’emplacement spécifié dans la banque d’informations Exchange.|
|[CreateItem Operation](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|Crée les éléments spécifiés dans la banque d’informations Exchange.|
|[FindConversation Operation](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|Énumère une liste des conversations dans le dossier spécifié dans la banque d’informations Exchange.|
|[FindFolder Operation](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|Cherche les sous-dossiers d’un dossier donné et retourne un ensemble de propriétés qui décrit l’ensemble de sous-dossiers.|
|[FindItem Operation](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|Identifie les éléments situés dans un dossier donné dans la banque d’informations Exchange.|
|[GetConversationItems operation](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|Obtient un ou plusieurs ensembles d’éléments organisés en nœuds dans une conversation.|
|[GetFolder Operation](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|Obtient les propriétés spécifiées et le contenu des dossiers de la banque d’informations Exchange.|
|[GetItem Operation](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|Obtient les propriétés spécifiées et le contenu des éléments de la banque d’informations Exchange.|
|[MarkAsJunk Operation](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|Déplace les messages électroniques vers le dossier Courrier indésirable, et ajoute ou supprime les expéditeurs des messages de la liste des expéditeurs bloqués.|
|[MoveItem Operation](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|Déplace les éléments dans un dossier de destination unique dans la banque d’informations Exchange.|
|[SendItem Operation](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|Envoie les messages électroniques situés dans la banque d’informations Exchange.|
|[Opération UpdateFolder](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|Modifie les propriétés des dossiers existants dans la banque d’informations Exchange.|
|[UpdateItem Operation](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|Modifie les propriétés des éléments existants dans la banque d’informations Exchange.|

## Authentification et autorisations pour la méthode makeEwsRequestAsync


Avec la méthode  **makeEwsRequestAsync**, la demande est authentifiée au moyen des informations d’identification du compte de messagerie de l’utilisateur actuel. La méthode  **makeEwsRequestAsync** gère ces informations d’identification pour vous, ce qui vous évite d’en fournir avec votre demande.


 >
  **Remarque**  L’administrateur du serveur doit utiliser l’applet de commande [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx) ou l’applet de commande [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx) pour définir le paramètre _OAuthAuthentication_ sur **true** sur le répertoire EWS du serveur d’accès client afin d’activer la méthode **makeEwsRequestAsync** pour effectuer des demandes EWS.

Votre complément doit spécifier l’autorisation **ReadWriteMailbox** dans son manifeste pour utiliser la méthode **makeEwsRequestAsync**. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox**, voir la section [Autorisation ReadWriteMailbox](../outlook/understanding-outlook-add-in-permissions.md#readwritemailbox-permission) dans [Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur](../outlook/understanding-outlook-add-in-permissions.md).


## Ressources supplémentaires



- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Confidentialité et sécurité pour les compléments Office](../../docs/develop/privacy-and-security.md)
    
- [Résolutions des limites de stratégie d’origine identique dans les compléments Office](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- [Référence EWS pour Exchange](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- [Applications de messagerie pour Outlook et EWS dans Exchange](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
Consultez la rubrique suivante pour créer des services principaux pour les compléments à l’aide de l’API Web ASP.NET :


- [Créer un service web pour un complément Office à l’aide de l’API Web ASP.NET](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [Principes fondamentaux de la création d’un service HTTP à l’aide de l’API Web ASP.NET](http://www.asp.net/web-api)
    
