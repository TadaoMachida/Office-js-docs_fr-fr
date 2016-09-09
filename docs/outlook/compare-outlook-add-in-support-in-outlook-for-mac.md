
# Comparaison de la prise en charge de compléments Outlook dans Outlook pour Mac et dans d’autres hôtes Outlook

Vous pouvez créer et exécuter un complément Outlook de la même façon dans Outlook pour Mac que dans les autres hôtes, y compris Outlook pour Windows, OWA pour périphériques et Outlook Web App, sans personnaliser JavaScript pour chaque hôte. Les appels du complément vers l’Interface API JavaScript pour Office fonctionnent généralement de la même manière, sauf pour les domaines décrits dans le tableau suivant.

 >**Remarque**  Outlook pour Mac prend en charge Interface API JavaScript pour Office dans Outlook en mode lecture uniquement.

|**Catégorie**|**Outlook pour Windows, OWA pour périphériques, Outlook Web App**|**Outlook pour Mac**|
|:-----|:-----|:-----|
|Versions d’office.js et du schéma de manifeste des Compléments Office pris en charge|Toutes les API dans Office.js et le schéma version 1.1.|<ul><li>Seules les API qui sont applicables en mode lecture. Un complément qui utilise les API nouvelles et étendues dans la version 1.1 d’office.js peut être activé, mais les API en mode composition ne s’exécutent pas correctement sur Outlook pour Mac. </li><li>Version 1.1 du schéma</li></ul>|
|Instances d’une série de rendez-vous périodiques|<ul><li>Peut obtenir l’ID d’élément et d’autres propriétés d’un rendez-vous principal ou d’une instance de rendez-vous d’une série périodique. </li><li>peut utiliser [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md#displayappointmentformitemid) pour afficher une instance ou le masque d’une série périodique.</li></ul>|<ul><li>Peut obtenir l’ID d’élément et d’autres propriétés du rendez-vous principal, mais pas ceux d’une instance d’une série périodique.</li><li>Peut afficher le rendez-vous principal d’une série périodique. Sans l’ID d’élément, ne peut pas afficher une instance d’une série périodique.</li></ul>|
|Type de destinataire d’un participant de rendez-vous|Peut utiliser [EmailAddressDetails.recipientType](../../reference/outlook/simple-types.md) pour identifier le type de destinataire d’un participant.|**EmailAddressDetails.recipientType** renvoie **undefined** pour les participants au rendez-vous.|
|Chaîne de version de l’hôte |Le format de la chaîne de version renvoyée par [diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) dépend du type de l’hôte. Par exemple :<ul><li>Outlook pour Windows : 15.0.4454.1002</li><li>Outlook Web App : 15.0.918.2</li></ul>|Exemple de la chaîne de version renvoyée par  **Diagnostics.hostVersion** sur Outlook pour Mac : 15.0 (140325)|
|Propriétés personnalisées d’un élément|Si le réseau tombe en panne, un complément peut toujours accéder aux propriétés personnalisées mises en cache.|Comme Outlook pour Mac ne met pas en cache les propriétés personnalisées, si le réseau tombe en panne, les compléments ne pourront pas y accéder.|
|Détails des pièces jointes|Le type de contenu et le nom des pièces jointes figurant dans un objet [AttachmentDetails](../../reference/outlook/Office.context.mailbox.md) dépendent du type d’hôte :<ul><li>Exemple JSON de <b>AttachmentDetails.contentType</b> : <b>"contentType": "image/x-png"</b>. </li><li><b>AttachmentDetails.name</b> ne contient aucune extension de nom de fichier. Par exemple, si la pièce jointe est un message dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de la pièce jointe serait <b>"name": "RE: Summer activity"</b>.</li></ul>|<ul><li>Exemple JSON de <b>AttachmentDetails.contentType</b>: <b>"contentType": "image/png"</b></li><li><b>AttachmentDetails.name</b> inclut toujours une extension de nom de fichier. Les pièces jointes qui sont des éléments de messagerie ont une extension .eml et les rendez-vous ont une extension .ics. Par exemple, si une pièce jointe est un message électronique dont l’objet est « RE: Summer activity », l’objet JSON qui représente le nom de pièce jointe sera <b>"name": "RE: Summer activity.eml".</b></li></ul>|
|Chaîne représentant le fuseau horaire dans les propriétés  **dateTimeCreated** et **dateTimeModified**|Par exemple : Jeu 13 mar 2014 14:09:11 GMT + 0800 (heure standard de la Chine)|Par exemple : Jeu 13 mar 2014 14:09:11 GMT + 0800 (CST)|
|Précision horaire de  **dateTimeCreated** et **dateTimeModified**|Si un complément utilise le code suivant, la précision est de l’ordre de la milliseconde.<br/><pre lang="javascript">JSON.stringify (Office.context.mailbox.item, null, 4) ;</pre>|La précision peut seulement atteindre une seconde.|

## Ressources supplémentaires



- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md)
    
