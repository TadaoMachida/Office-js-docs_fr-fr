
# Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office

Parfois, vos utilisateurs peuvent rencontrer des problèmes avec les compléments Office que vous développez. Par exemple, il se peut qu’un complément ne se charge pas ou soit inaccessible. Utilisez les informations de cet article pour résoudre les problèmes courants que vos utilisateurs rencontrent avec votre complément Office. 

Vous pouvez également utiliser [Fiddler](http://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.

Une fois le problème de l’utilisateur résolu, vous pouvez [répondre directement aux avis des clients dans l’Office Store](https://msdn.microsoft.com/library/jj635874.aspx).

## Erreurs courantes et étapes de dépannage

Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.



|**Message d’erreur**|**Solution**|
|:-----|:-----|
|Erreur d’application : impossible d’accéder au catalogue|Vérifiez les paramètres de pare-feu.Le « catalogue » se réfère à l’Office Store. Ce message indique que l’utilisateur ne peut pas accéder à l’Office Store.|
|Erreur d’application : cette application n’a pas pu être démarrée. Fermez cette boîte de dialogue pour ignorer le problème, ou cliquez sur « Redémarrer » pour réessayer.|Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/en-us/kb/2986156/).|
|Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »|Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité. Accédez à Outils >  **Paramètres d’affichage de compatibilité**.|
|Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.|Assurez-vous que le navigateur prend en charge le stockage local HTML5 ou réinitialisez les paramètres d’Internet Explorer.Pour plus d’informations sur les navigateurs pris en charge, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).|

## §LTA Le complément Outlook ne fonctionne pas correctement

§LTA Si un complément Outlook s’exécutant sous Windows ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer. 


- Accédez à Outils >  **Options Internet** > **Avancées**.
    
- Sous  **Parcourir**, décochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.
    
Nous vous recommandons de décocher ces paramètres uniquement pour résoudre le problème. Si vous ne les réactivez pas, vous recevrez des invites. Une fois que le problème est résolu, recochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.


## Le complément ne s’active pas dans Office 2013

Le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes :


1. connexion à son compte Microsoft dans Office 2013 ;
    
2. activation de la vérification à deux étapes pour son compte Microsoft ;
    
3. vérification de son identité après invitation lorsqu’il tente d’insérer un complément.
    
Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/en-us/kb/2986156/).


## Ressources supplémentaires



- [Débogage de compléments dans Office Online](../testing/debug-add-ins-in-office-online.md)
    
- [Charger une version test d’un complément Office sur iPad ou Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Débogage des compléments Office sur iPad et Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [Créer et déboguer des compléments Office dans Visual Studio](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md)
    
