
# Compléments Outlook avec volet personnalisé

Un volet personnalisé est un point d’extension d’un complément qui s’active lorsque certaines conditions spécifiques sont remplies par l’élément sélectionné. Il est défini dans l’élément  **VersionOverrides** manifeste du complément, avec toutes les commandes de complément implémentées par le complément. Pour plus d’informations, voir [Définir des commandes de complément dans votre manifeste de complément Outlook](../outlook/manifests/define-add-in-commands.md). Un volet personnalisé ne peut apparaître que dans une vue de message lu ou de participant à un rendez-vous. Il affiche une entrée dans la barre de complément. Lorsque l’utilisateur clique sur l’entrée, le volet personnalisé s’affiche dans le sens horizontal, au-dessus du corps de l’élément. L’affichage et le comportement sont identiques à ceux des compléments en mode Lecture qui n’implémentent pas de commandes de complément.

**Complément avec volet personnalisé en mode Lecture**

![Affiche un volet personnalisé dans un formulaire de lecture de message.](../../images/c585ab0a-6c33-42d0-a20f-5deb8b54f480.png)

L’exemple suivant définit un volet personnalisé pour des éléments qui sont des messages, qui ont une pièce jointe ou qui incluent une adresse. 



```
<ExtensionPoint xsi:type="CustomPane">
  <RequestedHeight>100< /RequestedHeight> 
  <SourceLocation resid="residReadTaskpaneUrl"/>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasAttachment"/>
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
  </Rule>
</ExtensionPoint>
```



-  **RequestedHeight** indique la hauteur voulue, en pixels, du complément de messagerie lorsqu’il est exécuté sur un ordinateur de bureau. Sinon, il est ignoré. Sa valeur peut être comprise entre 32 et 450. Si ce paramètre n’est pas défini, la valeur par défaut est de 350 px. Facultatif.
    
-  **SourceLocation** spécifie la page HTML qui fournit l’interface utilisateur du volet personnalisé. L’attribut **resid** est défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Resources**. Obligatoire.
    
-  
  **Rule** spécifie la règle ou l’ensemble de règles qui précisent dans quelles conditions le complément est activé. Ce paramètre est tel que défini dans [Manifestes des compléments Outlook](../outlook/manifests/manifests.md), sauf que la règle [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) a été modifiée de la manière suivante : l’élément **ItemType** est « Message » ou « AppointmentAttendee », et l’attribut **FormType** est absent. Pour plus d’informations, voir [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md).
    

## Ressources supplémentaires



- [Prise en main des compléments Outlook pour Office 365](https://dev.outlook.com/MailAppsGettingStarted)
    
- [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md)
    
- [Manifestes des compléments Outlook](../outlook/manifests/manifests.md)
    
- [Définir des commandes de complément dans votre manifeste de complément Outlook](../outlook/manifests/define-add-in-commands.md)
    
