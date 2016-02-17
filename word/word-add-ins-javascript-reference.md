# Référence des API JavaScript pour les compléments Word 

Recherchez des références d’API JavaScript pour les compléments Word.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Dans cette section

Voici les objets principaux de l’API JavaScript pour Word.

* [Body](word-add-ins-javascript-reference/body.md) : représente le corps d’un document ou d’une section.
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md): conteneur de contenu. Il s’agit d’une zone d’un document délimitée par des bordures et pouvant porter une étiquette qui sert à contenir un certain type de contenu. Par exemple, les objets ContentControl peuvent contenir des paragraphes de texte mis en forme et d’autres contrôles de contenu. Vous pouvez accéder à un objet ContentControl via la collection de contrôles de contenu du document, le corps du document, un paragraphe, une plage ou via un autre contrôle de contenu.
* [Document](word-add-ins-javascript-reference/document.md) : l’objet de niveau supérieur. Un objet Document comporte des [sections](word-add-ins-javascript-reference/section.md), un corps dans lequel se trouve le contenu du document et des informations d’en-tête/de pied de page.
* [Font](word-add-ins-javascript-reference/font.md) : permet de mettre en forme le texte d’un corps, d’un contrôle de contenu, d’un paragraphe ou d’une plage.
* [Image](word-add-ins-javascript-reference/inlinepicture.md) : représente une image incluse ancrée à un paragraphe.
* [Paragraph](word-add-ins-javascript-reference/paragraph.md) : représente un paragraphe unique d’une sélection, d’une plage ou d’un document. Vous pouvez accéder à un paragraphe via la collection de paragraphes d’un document, d’une plage ou d’une sélection. 
* [Range](word-add-ins-javascript-reference/range.md) : Représente une zone contiguë dans un document. Vous obtenez un objet de plage lorsque vous effectuez une sélection, que vous insérez du contenu dans le corps, dans un contrôle de contenu ou dans un paragraphe, ou lorsque vous obtenez un résultat de recherche. Vous pouvez définir et manipuler une plage sans modifier la sélection.
* [Section](word-add-ins-javascript-reference/section.md) :  définit différents en-têtes et pieds de page, ainsi que différentes configurations de mise en page pour un document. Vous pouvez accéder aux sections à partir de l’objet Document. 
* [Selection](word-add-ins-javascript-reference/document.md#getselection) : l’objet Sélection vous donne accès à la sélection de l’utilisateur dans le document ou au point d’insertion actif si aucun élément n’est sélectionné.

## Donnez-nous votre avis.

Votre avis compte beaucoup pour nous. 

* Consultez les documents et signalez-nous toute question ou tout problème à leur propos en [soumettant une question](https://github.com/OfficeDev/office-js-docs/issues) directement dans ce référentiel.
* Faites-nous part de vos expériences de programmation, de ce que vous souhaiteriez voir dans les futures versions, de vos questions sur les exemples de code, etc. Passez par [ce site](http://officespdev.uservoice.com/) pour soumettre vos suggestions et vos idées.

## Ressources supplémentaires

* [Compléments Word](word-add-ins.md)
* [Guide de programmation des compléments Word](word-add-ins-programming-guide.md)
* [Compléments Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Commencer à utiliser les compléments Office](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;Compléments Word sur GitHub&lt;/a&gt;
* [Explorateur d’extraits de code pour Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
