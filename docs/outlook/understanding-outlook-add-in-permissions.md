
# Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur

Les compléments Outlook indiquent le niveau d’autorisation requis dans leur manifeste. Les niveaux disponibles sont  **Restricted**,  **ReadItem**,  **ReadWriteItem** ou **ReadWriteMailbox**. Ces niveaux d’autorisation sont cumulatifs :  **Restreint** est le niveau le plus bas et chaque niveau supérieur inclut les autorisations de tous les niveaux inférieurs. L’autorisation **Lire/écrire dans la boîte aux lettres** comprend toutes les autorisations prises en charge.

Vous pouvez voir les autorisations demandées par un complément de messagerie avant de l’installer depuis l’Office Store. Vous pouvez également voir les autorisations requises des compléments installés dans le Centre d’administration Exchange.


## Autorisation restreint


L’autorisation  **Restreint** est le niveau d’autorisation le plus élémentaire. Vous pouvez indiquer **Restricted** dans l’élément [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) du manifeste pour demander cette autorisation. Outlook affecte par défaut cette autorisation à un complément de messagerie si celui-ci ne demande pas d’autorisation spécifique dans son manifeste.


### Vous pouvez :


- [Obtenir uniquement des entités spécifiques](../outlook/match-strings-in-an-item-as-well-known-entities.md) (numéro de téléphone, adresse, URL) de l’objet ou du corps de l’élément.
    
- Spécifier une [règle d’activation ItemIs](../outlook/manifests/activation-rules.md#itemis-rule) qui exige que l’élément actuel soit un type d’élément spécifique dans un formulaire de lecture ou de composition, ou une [règle ItemHasKnownEntity](../outlook/match-strings-in-an-item-as-well-known-entities.md) qui correspond à l’un des sous-ensembles plus petits d’entités connues prises en charge (numéro de téléphone, adresse, URL) dans l’élément sélectionné.
    
- Accéder aux propriétés et méthodes qui ne sont  **pas** associées aux informations spécifiques concernant l’utilisateur ou l’élément. (Consulter la section suivante pour obtenir la liste des membres qui le sont.)
    

### Vous ne pouvez pas :


- Utiliser une règle [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) sur l’entité de contact, d’adresse de messagerie, de suggestion de réunion ou de suggestion de tâche.
    
- Utiliser la règle [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) ou [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx).
    
- Accéder aux membres de la liste suivante qui appartiennent aux informations de l’utilisateur ou de l’élément. La tentative d’accès à des membres de cette liste renverra  **null** et entraînera un message d’erreur qui indiquera qu’Outlook exige une élévation des autorisations du complément de messagerie.
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) et tous ses membres enfants
    
  - [Location](../../reference/outlook/Location.md) et tous ses membres enfants
    
  - [Recipients](../../reference/outlook/Recipients.md) et tous ses membres enfants
    
  - [Subject](../../reference/outlook/Subject.md) et tous ses membres enfants
    
  - [Time](../../reference/outlook/Time.md) et tous ses membres enfants
    

## Autorisation ReadItem


L’autorisation  **ReadItem** est le niveau suivant d’autorisation dans le modèle d’autorisations. Vous pouvez indiquer **ReadItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation.


### Vous pouvez :


- [Lire toutes les propriétés](../outlook/item-data.md) de l’élément actuel dans un formulaire de lecture ou de [composition](../outlook/get-and-set-item-data-in-a-compose-form.md), par exemple, [item.to](../../reference/outlook/Office.context.mailbox.item.md) dans un formulaire de lecture et [item.to.getAsync](../../reference/outlook/Recipients.md) dans un formulaire de composition.
    
- [Obtenir un jeton de rappel pour obtenir des pièces jointes d’éléments](../outlook/get-attachments-of-an-outlook-item.md) ou l’élément complet.
    
- [Écrire des propriétés personnalisées](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx) définies par le complément sur cet élément.
    
- [Obtenir toutes les entités existantes connues](../outlook/match-strings-in-an-item-as-well-known-entities.md), et pas seulement un sous-ensemble, à partir de l’objet ou du corps de l’élément.
    
- Utiliser toutes les [entités connues](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) dans les règles [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) ou les [expressions régulières](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) dans les règles [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx). L’exemple suivant suit le schéma version 1.1. Il montre une règle qui active le complément si une ou plusieurs entités connues sont trouvées dans l’objet ou le corps du message sélectionné :
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### Vous ne pouvez pas :

Accéder à  **mailbox.makeEWSRequestAsync** ou aux méthodes d’écriture suivantes :


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## Autorisation ReadWriteItem


Vous pouvez indiquer  **ReadWriteItem** dans l’élément **Permissions** du manifeste pour demander cette autorisation. Les compléments de messagerie activés dans des formulaires de composition et utilisant des méthodes d’écriture (par exemple, **Message.to.addAsync** ou **Message.to.setAsync**) doivent utiliser au moins ce niveau d’autorisation.


### Vous pouvez :


- [Lire et écrire toutes les propriétés au niveau de l’élément](../outlook/item-data.md) concernant l’élément affiché ou en cours de composition dans Outlook.
    
- [Ajouter ou supprimer des pièces jointes](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md) de cet élément.
    
- Utiliser tous les autres membres de l’API JavaScript pour Office applicables aux compléments de messagerie, excepté  **Mailbox.makeEWSRequestAsync**.
    

### Vous ne pouvez pas :

Utiliser  **Mailbox.makeEWSRequestAsync**.


## Autorisation ReadWriteMailbox


L’autorisation  **ReadWriteMailbox** est le niveau d’autorisation le plus élevé dans le modèle d’autorisations. Vous pouvez indiquer **ReadWriteMailbox** dans l’élément **Permissions** du manifeste pour demander cette autorisation.

En plus de ce que prend en charge l’autorisation  **ReadWriteItem**, en utilisant  **Mailbox.makeEWSRequestAsync**, vous pouvez accéder aux opérations des services web Exchange (EWS) prises en charge afin d’effectuer les actions suivantes :


- Lire et écrire toutes les propriétés d’un élément de la boîte aux lettres de l’utilisateur.
    
- Créer, lire et écrire dans tous les dossiers ou tous les éléments de cette boîte aux lettres.
    
- Envoyer un élément depuis cette boîte aux lettres.
    
Grâce à  **mailbox.makeEWSRequestAsync**, vous pouvez accéder aux opérations des services web Exchange suivantes :


- [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
Toute tentative d’utilisation d’une opération non prise en charge entraînera une réponse d’erreur.


## Ressources supplémentaires



- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
