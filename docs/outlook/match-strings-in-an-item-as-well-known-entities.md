

# Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues


Avant d’envoyer un élément de message ou de demande de réunion, Exchange Server analyse le contenu de l’élément, identifie et marque certaines chaînes dans l’objet et le corps similaires aux entités connues d’Exchange (par exemple, adresses e-mail, numéros de téléphone, URL). Les demandes de réunion et les messages sont envoyés par Exchange Server dans une boîte de réception Outlook avec les entités connues marquées. 

À l’aide de l’interface API JavaScript pour Office, vous pouvez obtenir les chaînes correspondant à des entités connues spécifiques en vue de leur traitement ultérieur. Vous pouvez également spécifier une entité connue dans une règle du manifeste du complément pour qu’Outlook puisse activer votre complément quand l’utilisateur affiche un élément contenant des correspondances pour cette entité. Vous pouvez extraire et effectuer une action sur des correspondances pour cette entité. 

Pouvoir identifier ou extraire de telles instances à partir d’un message ou d’un rendez-vous sélectionné s’avère très pratique. Par exemple, vous pouvez créer un service de recherche téléphonique inversée en tant que complément Outlook. Le complément peut extraire des chaînes dans l’objet ou le corps de l’élément qui ressemblent à un numéro de téléphone, effectuer une recherche inversée et afficher le propriétaire enregistré de chaque numéro de téléphone.

Cette rubrique présente ces entités connues, montre des exemples de règles d’activation en fonction de ces entités et explique comment extraire des correspondances d’entités indépendamment de l’utilisation d’entités dans les règles d’activation.


## Prise en charge des entités connues


Exchange Server marque les entités connues dans un élément de message ou de demande de réunion après que l’élément a été envoyé par l’expéditeur et avant qu’il soit remis au destinataire. Ainsi, seuls les éléments ayant transité via Exchange sont marqués, et Outlook peut activer des compléments en fonction de ces marquages quand l’utilisateur affiche ces éléments. En revanche, quand l’utilisateur compose ou affiche un élément du dossier Éléments envoyés, Outlook ne peut pas activer les compléments en fonction des entités connues car l’élément n’a pas transité via Exchange. 

De même, vous ne pouvez pas extraire les entités connues dans les éléments en cours de composition ou situés dans le dossier Éléments envoyés, car ces éléments n’ont pas transité via Exchange et ne sont pas marqués. Pour plus d’informations sur les types d’éléments qui prennent en charge l’activation, voir [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins).

Le tableau suivant répertorie les entités qu’Exchange Server et Outlook prennent en charge et reconnaissent (d’où le nom « entités connues »), et le type d’objet d’une instance de chaque entité. La reconnaissance du langage naturel d’une chaîne en tant que l’une de ces entités est fondée sur un modèle d’apprentissage qui a été testé sur une grande quantité de données. Par conséquent, la reconnaissance n’est pas déterministe. Pour plus d’informations sur les conditions de reconnaissance, voir [Conseils d’utilisation des entités connues](#conseils-dutilisation-des-entités-connues).

 **Tableau 1. Entités prises en charge et leurs types**



|**Type d’entité**|**Conditions de reconnaissance**|**Type d’objet**|
|:-----|:-----|:-----|
|**Address**|Adresses aux États-Unis ; par exemple : 1234 Main Street, Redmond, WA 07722.Généralement, pour qu’une adresse puisse être reconnue, elle doit obéir à la structure d’une adresse postale des États-Unis, où la plupart des éléments sont présents, à savoir numéro de rue, nom de rue, ville, État et code postal. L’adresse peut être spécifiée sur une ou plusieurs lignes.|Objet JavaScript  **String**|
|**Contact**|Une référence aux informations d’une personne telles que reconnue en langage naturel.La reconnaissance d’un contact dépend du contexte. Par exemple, une signature à la fin d’un message ou le nom d’une personne apparaissant à proximité des informations suivantes : un numéro de téléphone, une adresse, une adresse électronique et une URL.|Objet [Contact](../../reference/outlook/simple-types.md)|
|**EmailAddress**|Adresses électroniques SMTP.|Objet JavaScript  **String**|
|**MeetingSuggestion**|Une référence à un événement ou une réunion. Par exemple, Exchange 2013 reconnaîtrait le texte suivant comme une suggestion de réunion :  _On se voit demain pour déjeuner ?_|Objet [MeetingSuggestion](../../reference/outlook/simple-types.md)|
|**PhoneNumber**|Numéros de téléphone des États-Unis ; par exemple :  _(235) 555-0110_|Objet [PhoneNumber](../../reference/outlook/simple-types.md)|
|**TaskSuggestion**|Phrases appelant une action. Par exemple :  _Veuillez mettre à jour la feuille de calcul._|Objet [TaskSuggestion](../../reference/outlook/simple-types.md)|
|**Url**|Adresse web qui spécifie explicitement l’identificateur et l’emplacement réseau d’une ressource web. Exchange Server n’exige pas le protocole d’accès dans l’adresse web et ne reconnaît pas les URL qui sont incorporées dans le texte du lien en tant qu’instances de l’entité  **Url**. Exchange Server peut correspondre aux exemples suivants : _www.youtube.com/user/officevideos_ _http://www.youtube.com/user/officevideos_|Objet JavaScript  **String**|
La figure 1 décrit comment Exchange Server et Outlook prennent en charge les entités connues pour les compléments et indique ce que les compléments peuvent faire avec ces entités connues. Voir [Récupération d’entités dans votre complément](#récupération-dentités-dans-votre-complément) et [Activation d’un complément basé sur l’existence d’une entité](#activation-dun-complément-basé-sur-lexistence-dune-entité) pour plus de détails sur l’utilisation de ces entités.


**Figure 1. Prise en charge des entités connues par Exchange Server, Outlook et les compléments**

![Prise en charge et utilisation des entités connues dans une application de messagerie](../../images/mod_off15_mailapp_wellknownentities_curvedlines.png)


## Autorisations d’extraction d’entités


Pour extraire les entités de votre code JavaScript ou pour activer votre complément à partir de l’existence de certaines entités connues, assurez-vous que vous avez demandé les autorisations appropriées dans le manifeste du complément.

La spécification de l’autorisation restreinte permet à votre complément d’extraire l’entité  **Address**,  **MeetingSuggestion** ou **TaskSuggestion**. Pour extraire l’une des autres entités, spécifiez l’autorisation Élément en lecture, Lecture/écriture ou Lecture/écriture de boîte aux lettres. Dans le manifeste, utilisez l’élément [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) et indiquez l’autorisation appropriée : **Restricted**,  **ReadItem**,  **ReadWriteItem** ou **ReadWriteMailbox**, comme dans l’exemple suivant :




```XML
<Permissions>ReadItem</Permissions>
```


## Récupération d’entités dans votre complément


Tant que l’objet ou le corps de l’élément consulté par l’utilisateur contient des chaînes qu’Exchange et Outlook peuvent reconnaître comme des entités connues, ces instances sont disponibles pour les compléments, et ce même si un complément n’est pas activé en fonction des entités connues. Avec les autorisations appropriées, vous pouvez utiliser la méthode  **getEntities** ou **getEntitiesByType** pour récupérer les entités connues qui sont présentes dans le message ou le rendez-vous actuel. La méthode **getEntities** renvoie un tableau d’objets [Entities](../../reference/outlook/simple-types.md), qui contient toutes les entités connues de l’élément. Si vous êtes intéressé par un type particulier d’entités, utilisez la méthode  **getEntitiesByType**, qui renvoie uniquement un tableau des entités souhaitées. L’énumération [EntityType](../../reference/outlook/Office.MailboxEnums.md) représente tous les types d’entités connues que vous pouvez extraire.

Après l’appel de  **getEntities**, vous pouvez utiliser la propriété correspondante de l’objet  **Entities** pour obtenir un tableau des instances d’un type d’entité. En fonction du type d’entité, les instances du tableau peuvent être uniquement des chaînes ou correspondre à des objets spécifiques. Dans l’exemple de la figure 1, pour obtenir les adresses de l’élément, accédez au tableau renvoyé par `getEntities().addresses[]`. La propriété  **Entities.addresses** renvoie un tableau de chaînes qu’Outlook reconnaît comme étant des adresses postales. De même, la propriété **Entities.contacts** renvoie un tableau d’objets **Contact** qu’Outlook reconnaît comme étant des coordonnées. Le tableau 1 répertorie le type d’objet d’une instance de chaque entité prise en charge.

L’exemple suivant illustre comment récupérer des adresses trouvées dans un message.




```
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities &amp;&amp; null != entities.addresses &amp;&amp; undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## Activation d’un complément basé sur l’existence d’une entité


Une autre façon d’utiliser des entités connues consiste à faire en sorte qu’Outlook active votre complément selon l’existence de types d’entités dans l’objet ou le corps de l’élément actuellement affiché. Pour cela, spécifiez une règle  **ItemHasKnownEntity** dans le manifeste du complément. Le type simple [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) représente les différents types d’entités connues pris en charge par les règles **ItemHasKnownEntity**. Une fois votre complément activé, vous pouvez également récupérer les instances de ces entités pour répondre à vos besoins, comme le décrit la section précédente [Récupération d’entités dans votre complément](#récupération-dentités-dans-votre-complément). 

Vous pouvez éventuellement appliquer une expression régulière dans une règle  **ItemHasKnownEntity**, de façon à filtrer des instances d’une entité et à faire en sorte qu’Outlook active un complément uniquement sur un sous-ensemble des instances de l’entité. Par exemple, vous pouvez spécifier un filtre pour l’entité d’adresse dans un message qui contient un code postal de l’État de Washington commençant par « 98 ». Pour appliquer un filtre sur les instances d’une entité, utilisez les attributs  **RegExFilter** et **FilterName** dans l’élément [Rule](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) du type [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx).

Comme avec d’autres règles d’activation, vous pouvez spécifier plusieurs règles afin de former une collection de règles pour votre complément. L’exemple suivant applique une opération AND sur deux règles : une règle  **ItemIs** et une règle **ItemHasKnownEntity**. Cette collection de règles active le complément lorsque l’élément en cours est un message et qu’Outlook reconnaît une adresse dans l’objet ou le corps de cet élément.




```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

L’exemple suivant utilise  **getEntitiesByType** pour l’élément en cours afin de définir une variable `addresses` pour les résultats de la collection de règles précédente.




```
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

L’exemple de règle  **ItemHasKnownEntity** suivant active le complément chaque fois qu’une URL se trouve dans l’objet ou le corps de l’élément actuel, et qu’elle contient la chaîne « youtube », indépendamment de la casse de cette chaîne.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

L’exemple suivant utilise  **getFilteredEntitiesByName(name)** de l’élément actuel pour définir une variable `videos` afin d’obtenir un tableau de résultats correspondant à l’expression régulière dans la règle **ItemHasKnownEntity** précédente.




```
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## Conseils d’utilisation des entités connues


Si vous utilisez des entités connues dans votre complément, vous devez connaître certaines informations et limites. Ce qui suit s’applique aussi longtemps que votre complément est activé, quand l’utilisateur lit un élément contenant des correspondances d’entités connues et indépendamment de l’utilisation ou non d’une règle  **ItemHasKnownEntity** :


1. Vous pouvez extraire des chaînes qui sont des entités connues uniquement si les chaînes sont en anglais.
    
2. Vous pouvez extraire des entités connues des 2 000 premiers caractères du corps de l’élément, mais pas au-delà. Cette limite de taille permet d’équilibrer le besoin de fonctionnalité et les performances, de sorte qu’Exchange Server et Outlook ne soient pas ralentis par l’analyse et l’identification des instances d’entités connues dans les longs messages et rendez-vous. Notez que cette limite est indépendante du fait que le complément indique une règle  **ItemHasKnownEntity**. Si le complément n’utilise pas cette règle, la limite de traitement de règle est celle décrite au point 2 ci-dessous pour les clients riches Outlook.
    
3. Vous pouvez extraire des entités à partir de rendez-vous, qui sont des réunions organisées par une personne autre que le propriétaire de la boîte aux lettres. Vous ne pouvez pas extraire d’entités à partir d’éléments de calendrier qui ne sont pas des réunions ou de réunions organisées par le propriétaire de la boîte aux lettres.
    
4. Vous pouvez extraire des entités du type  **MeetingSuggestion** uniquement à partir de messages et non de rendez-vous.
    
5. Vous pouvez extraire des URL qui existent de façon explicite dans le corps d’élément, mais pas des URL incorporées dans un texte de lien hypertexte du corps d’élément HTML. Pensez à utiliser une règle  **ItemHasRegularExpressionMatch** à la place pour obtenir à la fois des URL explicites et incorporées. Spécifiez **BodyAsHTML** en tant que _PropertyName_, ainsi qu’une expression régulière correspondant à des URL comme  _RegExValue_.
    
6. Vous ne pouvez pas extraire des entités à partir d’éléments dans le dossier Éléments envoyés.
    
En outre, les dispositions suivantes s’appliquent si vous utilisez une règle [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx), et cela peut avoir une incidence sur les scénarios pour lesquels vous souhaiteriez que votre complément soit activé :


1. Lors de l’utilisation de la règle  **ItemHasKnownEntity**, attendez-vous à ce qu’Outlook mette en correspondance uniquement les chaînes d’entité en anglais, quels que soient les paramètres régionaux par défaut spécifiés dans le manifeste.
    
2. Lorsque votre complément est en cours d’exécution sur un client riche Outlook, attendez-vous à ce qu’Outlook applique la règle  **ItemHasKnownEntity** sur le premier mégaoctet du corps de l’élément uniquement, et non au-delà de cette limite.
    
3. Vous ne pouvez pas utiliser de règle  **ItemHasKnownEntity** pour activer un complément pour les éléments du dossier Éléments envoyés.
    

## Ressources supplémentaires



- [Créer des compléments Outlook pour des formulaires de lecture](../outlook/read-scenario.md)
    
- [Extraire des chaînes d’entité d’un élément Outlook](../outlook/extract-entity-strings-from-an-item.md)
    
- [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md)
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Présentation des autorisations de complément Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    
