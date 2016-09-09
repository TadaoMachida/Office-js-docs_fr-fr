
# Débogage des compléments Office sur iPad et Mac

Vous pouvez utiliser Visual Studio pour le développement et le débogage des compléments sur Windows. Toutefois, vous ne pouvez pas l’utiliser pour déboguer les compléments sur iPad ou sur Mac. Dans la mesure où les compléments sont développés dans le code HTML et Javascript, ils devraient fonctionner sur différentes plateformes. Il peut toutefois exister de légères différences dans l’affichage du code HTML dans les différents navigateurs. Cette rubrique explique comment déboguer les compléments en exécution sur iPad ou sur Mac. 

## Débogage avec Vorlon.js 

Vorlon.js est un débogueur de pages web, semblable aux outils F12, conçu pour fonctionner à distance et pour vous permettre de déboguer des pages web sur différents appareils. Pour plus d’informations, accédez au [site web de Vorlon](http://www.vorlonjs.com).  

Pour installer et configurer Vorlon : 

1.  Installez [Node.js](https://nodejs.org) si ce n’est pas déjà fait. 

2.  Installez Vorlon à l’aide de npm avec la commande suivante : `sudo npm i -g vorlon` 

3.  Exécutez le serveur Vorlon avec la commande `vorlon`. 

4.  Ouvrez une fenêtre de navigateur et accédez à [http://localhost:1337](http://localhost:1337), qui correspond à l’interface Vorlon.

5.  Ajoutez la balise de script suivante à la section `<head>` du fichier home.html (ou fichier HTML principal) de votre complément :
```    
<script src="http://localhost:1337/vorlon.js"></script>    
```  

>**Remarque :** vous devez activer le protocole HTTPS dans Vorlon pour utiliser Vorlon.js afin de déboguer des compléments. Pour savoir comment procéder, voir le billet sur le [plug-in VorlonJS utilisé pour le débogage du complément Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).

Désormais, chaque fois que vous ouvrez le complément sur un appareil, il apparaît dans la liste Clients dans Vorlon (sur le côté gauche de l’interface Vorlon). Vous pouvez surligner à distance des éléments DOM, exécuter à distance des commandes et bien plus encore.  

![Capture d’écran de l’interface Vorlon.js](../../images/vorlon_interface.png)

Un plug-in Vorlon dédié pour les compléments Office ajoute des fonctionnalités supplémentaires, telles que l’interaction avec les API Office.js. Pour plus d’informations, voir le billet sur le [plug-in VorlonJS utilisé pour le débogage du complément Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/). Pour activer le plug-in des compléments Office : 

1.  Clonez localement la branche dev du référentiel GitHub Vorlon.js en utilisant les commandes suivantes : 
```
git clone https://github.com/MicrosoftDX/Vorlonjs.git
git checkout dev
npm install
```

2.  Ouvrez le fichier **config.json** situé dans /Vorlon/Server/config.json. Pour activer le plug-in du complément Office, définissez la propriété **enabled** sur **true**.

![Capture d’écran de la section Plugins de config.json](../../images/vorlon_plugins_config.png) 
