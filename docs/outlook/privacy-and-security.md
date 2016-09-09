
# Confidentialité, autorisations et sécurité pour les compléments Outlook
Les utilisateurs finaux, les développeurs et les administrateurs peuvent appliquer les niveaux d’autorisation hiérarchisés du modèle de sécurité pour les compléments Outlook afin de contrôler les performances et la confidentialité.



Cet article décrit les autorisations que les compléments Outlook peuvent demander, et examine le modèle de sécurité selon les perspectives suivantes :

- Office Store - intégrité des compléments.
    
- Utilisateurs finaux - préoccupations liées à la confidentialité et aux performances.
    
- Développeurs - choix d’autorisations et limites d’utilisation des ressources.
    
- Administrateurs - privilèges pour définir des seuils de performances.
    

## Modèle d’autorisations


Comme la façon dont les clients perçoivent la sécurité des compléments peut avoir une incidence sur l’adoption de ces derniers, la sécurité des compléments Outlook repose sur un modèle d’autorisations à plusieurs niveaux. Un complément Outlook indique le niveau d’autorisations dont il a besoin, identifiant ainsi l’accès dont il peut disposer et les actions qu’il peut effectuer sur les données de la boîte aux lettres du client. 

Le schéma de manifeste version 1.1 comprend quatre niveaux d’autorisation. 


**Tableau 1. Niveaux d’autorisation d'un complément**


|**Niveau d’autorisation**|**Valeur dans le manifeste du complément Outlook**|
|:-----|:-----|
|Restricted|Restreint|
|Lire l’élément|ReadItem|
|Lire/écrire dans l’élément|ReadWriteItem|
|Lire/écrire dans la boîte aux lettres|ReadWriteMailbox|
Les quatre niveaux d’autorisation sont cumulatifs : l’autorisation  **Lire/écrire dans la boîte aux lettres** inclut les autorisations **Lire/écrire dans l’élément**,  **Lire l’élément** et **Restreint**. L’autorisation  **Lire/écrire dans l’élément** inclut les autorisations **Lire l’élément** et **Restreint**. Enfin l’autorisation  **Lire l’élément** inclut l’autorisation **Restreint**. La figure 1 montre les quatre niveaux d’autorisation et décrit les possibilités offertes à l’utilisateur final, au développeur et à l’administrateur par chaque niveau. Pour plus d’informations sur ces autorisations, voir [Utilisateurs : problèmes de confidentialité et de performance](#utilisateurs-problèmes-de-confidentialité-et-de-performance), [Développeurs : choix d’autorisations et limites d’utilisation des ressources.](#développeurs-choix-dautorisations-et-limites-dutilisation-des-ressources.) et [Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur](../outlook/understanding-outlook-add-in-permissions.md). 


**Figure 1. Association du modèle d’autorisation à quatre niveaux à l’utilisateur, au développeur et à l’administrateur**

![Modèle d’autorisations à 4 niveaux pour le schéma d’applications de messagerie v1.1](../../images/olowa15wecon_Permissions_4Tier.png)


## Office Store : intégrité des compléments


Le Office Store héberge des compléments pouvant être installés par les utilisateurs finals et les administrateurs. Le Office Store applique les mesures suivantes pour maintenir l’intégrité de ces compléments Outlook :


- Oblige le serveur hôte d’un complément à toujours utiliser SSL (Secure Socket Layer) pour communiquer.
    
- Oblige un développeur à fournir une preuve d’identité, un accord contractuel et une politique de confidentialité conforme pour soumettre les compléments. 
    
- Archive les compléments en mode lecture seule.
    
- Prend en charge un système d’évaluation par les utilisateurs pour les compléments disponibles afin de promouvoir une communauté exerçant une auto surveillance.
    

## Utilisateurs : problèmes de confidentialité et de performance


Le modèle de sécurité résout les problèmes de sécurité, de confidentialité et de performance des utilisateurs des manières suivantes :


- Les messages des utilisateurs qui sont protégés par la Gestion des droits relatifs à l’information (IRM) d’Outlook n’ont pas d’interaction avec les compléments Outlook.
    
- Avant d’installer un complément de l’Office Store, les utilisateurs finals peuvent voir l’accès dont peut disposer le complément, ainsi que les actions qu’il peut effectuer sur leurs données, et doivent explicitement confirmer qu’ils veulent poursuivre. Aucun complément Outlook n’est automatiquement transmis sur un ordinateur client sans une validation manuelle par l’utilisateur ou l’administrateur.
    
- L’octroi de l’autorisation  **Restreint** permet au complément Outlook d’avoir un accès limité uniquement sur l’élément actuel. L’octroi de l’autorisation **Lire l’élément** permet au complément Outlook d’accéder à des informations d’identification personnelle, par exemple les noms et les adresses électroniques des expéditeurs et des destinataires, uniquement sur l’élément actuel.
    
- Un utilisateur final peut installer un complément Outlook uniquement pour lui-même. Les compléments de messagerie ayant une incidence sur l’organisation sont installés par un administrateur.
    
- Les utilisateurs peuvent installer des compléments Outlook qui activent des scénarios contextuels prisés par les utilisateurs tout en minimisant les risques de sécurité pour ces derniers.
    
- Les fichiers manifeste de compléments Outlook installés sont sécurisés dans le compte de messagerie de l’utilisateur.
    
- Les données échangées avec des serveurs hébergeant des Compléments Office sont toujours chiffrées conformément au protocole SSL (Secure Socket Layer).
    
- Applicable uniquement aux clients riches Outlook : les clients riches Outlook surveillent la performance des compléments Outlook installés, exercent un contrôle de gouvernance et désactivent les compléments Outlook qui dépassent les limites pour les aspects suivants :
    
      - Response time to activate
    
  - Nombre de défaillances d’activation ou de réactivation
    
  - Utilisation de la mémoire
    
  - Utilisation du processeur
    

    La gouvernance dissuade les attaques par déni de service et maintient les performances des compléments à un niveau raisonnable. La barre Entreprise indique aux utilisateurs les compléments Outlook que le client riche Outlook a désactivés sur la base d’un tel contrôle de gouvernance.
    
- À tout moment, les utilisateurs finals peuvent vérifier les autorisations demandées par les compléments Outlook installés, et désactiver ou activer ultérieurement tout complément Outlook dans le Centre d’administration Exchange.
    

## Développeurs : choix d’autorisations et limites d’utilisation des ressources.


Le modèle de sécurité fournit aux développeurs des niveaux précis d’autorisations à choisir, et de strictes directives de performance à observer.


### Les autorisations à plusieurs niveaux augmentent la transparence

Les développeurs doivent suivre le modèle d’autorisations à plusieurs niveaux pour assurer la transparence et apaiser les inquiétudes des utilisateurs concernant ce que les compléments peuvent faire à leurs données et leur boîte aux lettres, en faisant la promotion indirecte de l’adoption du complément :


- Les développeurs demandent un niveau approprié d’autorisation pour un complément Outlook en fonction de la manière dont il doit être activé, et de son besoin de lire ou d’écrire certaines propriétés d’un élément, ou de créer et d’envoyer un élément.
    
- Les développeurs demandent une autorisation en utilisant l’élément [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) dans le manifeste du complément Outlook, en affectant une valeur **Restricted**,  **ReadItem**,  **ReadWriteItem** ou **ReadWriteMailbox**, selon le cas. 
    
     >**Remarque** L’autorisation **ReadWriteItem** est disponible à partir du schéma de manifeste version 1.1.

    L’exemple suivant demande l’autorisation **Lire l’élément**.
    


```XML
  <Permissions>ReadItem</Permissions>
```

- Les développeurs peuvent demander l’autorisation  **Restreint** si le complément Outlook est activé lorsqu’un type spécifique d’élément Outlook (rendez-vous ou message) ou des entités extraites spécifiques (numéro de téléphone, adresse, URL) sont présents dans l’objet ou le corps de l’élément. Par exemple, la règle suivante active le complément Outlook si une ou plusieurs des trois entités (numéro de téléphone, adresse postale ou URL) se trouvent dans l’objet ou le corps du message actuel.
    
```XML
  <Permissions>Restricted</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
    </Rule>
</Rule>
```

- Les développeurs doivent demander l’autorisation  **Lire l’élément** si le complément Outlook doit lire les propriétés de l’élément actuel autres que les entités extraites par défaut, ou écrire des propriétés personnalisées définies par le complément sur l’élément actuel, mais ne nécessite pas de lire ou d’écrire d’autres éléments, ou de créer ou d’envoyer un message dans la boîte aux lettres de l’utilisateur. Par exemple, un développeur doit demander l’autorisation **Lire l’élément** si un complément Outlook doit rechercher une entité telle qu’une suggestion de réunion, une suggestion de tâche, une adresse électronique, ou un nom de contact dans l’objet ou le corps de l’élément, ou utilise une expression régulière pour s’activer.
    
- Les développeurs doivent demander l’autorisation  **Lire/écrire dans l’élément** si le complément Outlook doit écrire dans les propriétés de l’élément composé, comme les noms des destinataires, les adresses de messagerie, le corps et l’objet, ou s’il a besoin d’ajouter ou de supprimer des pièces jointes d’élément.
    
- Les développeurs demandent l’autorisation  **Lire/écrire dans la boîte aux lettres** uniquement si le complément Outlook doit effectuer une ou plusieurs des actions suivantes à l’aide de la méthode [mailbox.makeEWSRequestAsync](../../reference/outlook/Office.context.mailbox.md) :
    
      - Read or write to properties of items in the mailbox.
    
  - Créer, lire, écrire ou envoyer des éléments dans la boîte aux lettres.
    
  - Créer, lire ou écrire dans des dossiers de la boîte aux lettres.
    

### Réglage de l’utilisation des ressources

Les développeurs doivent connaître les limites de l’utilisation des ressources pour l’activation, incorporer le réglage des performances dans leur flux de travail de développement, afin de réduire le risque d’un complément peu performant refusant le service de l’hôte. Les développeurs doivent suivre les directives concernant la conception des règles d’activation telles que décrites dans [Limites d’activation et d’API JavaScript des compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Si un complément Outlook est destiné à être exécuté sur un client riche Outlook, les développeurs doivent vérifier que les performances du complément se situent dans les limites d’utilisation des ressources.


### Autres mesures visant à promouvoir la sécurité de l’utilisateur

Les développeurs doivent connaître et planifier les éléments suivants :


- Les développeurs ne peuvent pas utiliser de contrôles ActiveX dans les compléments car ils ne sont pas pris en charge.
    
- Les développeurs doivent effectuer ce qui suit lors de la soumission d’un complément Outlook à l’Office Store :
    
      - Produce an Extended Validation (EV) SSL certificate as a proof of identity.
    
  - Héberger le complément qu’ils soumettent sur un serveur web qui prend en charge SSL.
    
  - Produire une stratégie de confidentialité conforme.
    
  - Être prêts à signer un accord contractuel lors de la soumission du complément.
    

## Administrateurs : privilèges


Le modèle de sécurité fournit les droits et les responsabilités suivants aux administrateurs :


- Peut empêcher les utilisateurs d’installer un complément Outlook, notamment les compléments sur l’Office Store.
    
- Peut désactiver ou activer tout complément Outlook sur le Centre d’administration Exchange.
    
- Applicable uniquement à Outlook pour Windows : peut remplacer les paramètres de seuil de performance par des paramètres du Registre Objet de stratégie de groupe (GPO).
    


## Ressources supplémentaires



- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Confidentialité et sécurité pour les compléments Office](../../docs/develop/privacy-and-security.md)
    
- [API de complément Outlook](../outlook/apis.md)
    
- [Demande d’autorisations d’utilisation de l’API dans des compléments de contenu et de volet des tâches](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
    
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
