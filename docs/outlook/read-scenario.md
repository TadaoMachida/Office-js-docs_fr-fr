
# Créer des compléments Outlook pour des formulaires de lecture

Les compléments de lecture sont des compléments Outlook activés dans le volet de lecture ou l’inspecteur de lecture d’Outlook. Contrairement aux compléments de composition (qui sont des compléments Outlook activés lorsqu’un utilisateur crée un message ou un rendez-vous), les compléments de lecture sont disponibles dans les scénarios suivants :


- Affichage d’un message électronique, d’une demande de réunion, d’une réponse à une demande de réunion ou d’une annulation de réunion.*
    
- Affichage d’un élément de réunion dans lequel l’utilisateur est un participant.
    
- Affichage d’un élément de réunion dans lequel l’utilisateur est l’organisateur (version RTM d’Outlook 2013 et d’Exchange 2013 uniquement).
    
     >**Remarque**  À partir de la version Office 2013 SP1, si l’utilisateur visualise un élément de réunion dont il est l’organisateur, seuls les compléments de composition peuvent être activés et disponibles. Les compléments de lecture ne sont plus disponibles dans ce scénario.
* Outlook n’active pas les compléments dans un formulaire de lecture pour certains types de messages, y compris les éléments qui sont les pièces jointes d’un autre message et les éléments des dossiers Brouillons ou Courrier indésirable d’Outlook, ou encore ceux chiffrés ou protégés d’autres façons.

Dans chacun de ces scénarios de lecture, Outlook active les compléments lorsque leurs conditions d’activation sont respectées. Les utilisateurs peuvent ensuite choisir et ouvrir les compléments activés dans la barre de compléments du volet de lecture ou de l’inspecteur de lecture. La figure 1 montre le complément  **Bing Cartes** qui a été activé et ouvert alors que l’utilisateur lit un message contenant une adresse géographique.


**Figure 1. Volet de complément montrant le complément Bing Cartes en action pour le message Outlook sélectionné qui contient une adresse**

![Application de messagerie avec carte Bing dans Outlook](../../images/off15appsdk_BingMapMailAppScreenshot.jpg)


## Types de complément disponibles en mode de lecture


Les compléments de lecture peuvent correspondre à n’importe quelle combinaison des types suivants.


- [Commandes de complément pour Outlook](../outlook/add-in-commands-for-outlook.md)
    
- [Compléments Outlook contextuels](../outlook/contextual-outlook-add-ins.md)
    
- [Compléments Outlook de volet personnalisé](../outlook/custom-pane-outlook-add-ins.md)
    

## Fonctionnalités de l’API disponibles pour les compléments de lecture


Pour obtenir la liste des fonctionnalités que l’API JavaScript pour Office fournit aux compléments Outlook dans les formulaires de lecture, voir les tableaux 1 et 2 dans [Fonctionnalités des applications de messagerie par version](http://msdn.microsoft.com/library/f34e2f44-8c9d-4e90-b1d7-3f29506adb92%28Office.15%29.aspx). 

Voir aussi :


- Pour activer les compléments dans les formulaires de lecture, voir le tableau 1 dans : [Spécifier des règles d’activation dans un manifeste](../outlook/manifests/activation-rules.md#specify-activation-rules-in-a-manifest).
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Mettre en correspondance des chaînes dans un élément Outlook en tant qu’entités connues](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Extraire des chaînes d’entité d’un élément Outlook](../outlook/extract-entity-strings-from-an-item.md)
    
- [Obtenir des pièces jointes d’un élément Outlook à partir du serveur](../outlook/get-attachments-of-an-outlook-item.md)
    

## Ressources supplémentaires



- [Prise en main des compléments Outlook pour Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
