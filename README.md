# NegoApp
Logiciel simple de négociation bilatérale réalisé dans le cadre du Master en Relations Internationales de l'ULB, année 2017-2018.

//////////////ORGANISATION DU GITHUB//////////////

1° 
2° Excel
3° 


//////////////FONCTIONNEMENT DU PROGRAMME//////////////

1) Lancer l'exécutable ;

2) Via l'onglet Fichier, ouvrir un fichier Excell (.xlsx) pour la lecture par le programme des données (/!\ Étape obligatoire) ;

3) L'onglet Caractéristiques sert à vérifier les informations lues par le programme, ou à les modifier en validant par le bouton "Valider".
Quelques explications de bases pour tous ces paramètres.

- Numéro de l'agent : affiche les caractéristiques de l'agent sélectionné
- Coefficient d'affichage : tous les paramètres ont été codés pour être des nombres single entre 0 et 1, le coefficient d'affichage entré permet de comparer plus facilement le résultat du programme avec des données empiriques. A priori, pour n'importe quel autre test autre qu'empirique, il n'est pas utile de le modifier.
- Objectifs quantitatifs :
        o Minimum requis : on part du principe que les agents négocient dans un jeu à somme nulle, un gâteau de taille 1. Il s'agit du minimum en deçà duquel un agent ne peut absolument pas accepter de proposition.
        o Maximum espéré : idem, mais il s'agit du maximum. Il est parfois utilisé comme valeur de proposition, mais est assez peu utilisé par les agents de ce programme en dehors de l'objectif Partager (voir Objectif de l'agent).
        o Moyenne de l'Optimum : il s'agit du chiffre-clé autour duquel va se construire la grande majorité des stratégies de propositions des agents. C'est le paramètre-clé à faire varier pour les différents essais. Pour les négociations simples, le chiffre est pris comme une valeur absolue. Pour les négociations dynamiques (multiples), ce chiffre est pris comme une moyenne.
        o Écart-type de l'Optimum : c'est un paramètre de variabilité selon une loi normale de la valeur de l'optimum. C'est seulement utile pour les tests dit "dynamiques" (plusieurs tirages avec remise).
- Contraintes internes :
        o Aversion au risque : à partir du moment où la proposition faite se situe en dehors des limites d'acceptabilité de l'agent considéré (étendue min-max), plus sa valeur est faible (il n'y a pas d'aversion à risquer), plus il y a une chance que l'agent considéré quitte prématurément les rounds des négociations.
        o Aversion au regret : à partir du moment où les tours de négociations s'approchent de la limite fixée par l'utilisateur, plus sa valeur est forte, plus il y a une chance que l'agent considéré abandonne sa stratégie de base pour faire des propositions médianes, l'agent ne désirant pas regretter d'être passé à côté d'une occasion d'obtenir un accord.
        o Mandat de négociation : paramètre non-utilisé par le programme actuellement. Il est censé représenter la part négociable de l'agent sur le nombre de gâteaux de taille 1 à partager entre les négociateur (négociation multi-sujets).
- Variables psychologiques :
        o Degré de confiance : préalablement à la négociation, il y a un premier tour dit de "révélation" quant aux préférences de base des agents. Cette étape vise à vaincre l'incertitude qui entoure la négociation et les informations échangées ici vont être la base à partir de laquelle les agents vont réellement négocier entre eux. Ce paramètre dispose qu'entre 0 et 0,3, l'agent mentira (donnera son max plus l'amplitude au lieu de son optimum), entre 0,3 et 0,6, il bluffera (donnera son optimum plus l'amplitude), et à plus de 0,6, il dira la vérité (donnera son optimum). Ces données seront toujours visibles dans la fenêtre des négociations.
        o Objectif de l'Agent : Maximiser représente l'objectif de toute stratégie, où l'agent cherchera toujours à obtenir son maximum plutôt que son optimum. Partager représente une vision que l'on pourrait qualifier de plus bilatérale, où l'agent cherchera toujours à faire prévaloir l'intérêt commun des deux agents plutôt que son maximum individuel.
- Variabilité des données :
        o Amplitude des variations : il s'agit d'un paramètre permettant la variation des résultats des négociations. Dans le code actuel, un agent acceptera toute proposition faite par un autre si sa valeur est comprise entre son optimum +/- l'amplitude. Ce paramètre sert aussi comme base d'amplitude dans la phase de révélation, comme spécifié précédemment (voir Degré de confiance).
        o Degré de certitude : il s'agit d'un autre paramètre permettant la variation des résultats. Il représente l'incertitude des données transmises à l'agent le concernant lui pour démarrer la négociation. Dans l'état actuel du programme, si l'objectif d'un agent est de Maximiser, il partira du principe que ses Minimum, Maximum et Optimum étant incertains au degré spécifié ici, il en augmentera ses objectifs quantitatifs pour obtenir des valeurs globalement plus hautes. Si son objectif est de Partager, il va au contraire élargir sa palette d'acceptabilité en en diminuant son Minimum, augmentant son Maximum et en en conservant tel quel son Optimum. Cette étape est préalable à l'étape de la révélation des données, puis des négociations.

4) L'onglet Action de Négociation permet de choisir quel type de négociation l'on veut faire exécuter par le programme.

- Statique : une seule négociation possible. Le résultat de la révélation est détaillée. Le résultat de la négociation est détaillé au tour par tour. Il y a également un rappel des objectifs de chacun (modifiable ici par commodité), et un nouveau paramètre qui s'appelle stratégie.
        o Médiane : s'il n'y a pas de valeur préalable, l'agent propose son optimum. Sinon, il propose toujours la médiane entre la dernière valeur proposée et sa dernière valeur, jusqu'à l'obtention d'un accord. Problème de finition ici puisqu'il faudrait définir une borne en deçà de laquelle l'accord est réputé trouvé (en nombre de chiffres après la virgule par ex.).
        o Rigide : stratégie primaire où l'agent propose sa valeur de base à chaque tour, point. N'est intéressant qu'en conjonction avec d'autres stratégies.
        o Rigide avec risque : prise en compte du facteur risque (départ prématuré si valeur proposée est exagérément faible/haute)
        o Rigide avec regret : prise en compte du facteur regret, sans le risque (bifurcation vers une stratégie médiane si la fin des tours de négociation approche)
        o Rigide avec risque et regret : prise en compte des deux facteurs conjointement
- Dynamique : plusieurs négociations possibles. Le résultat de la révélation de chaque négociation est masqué. Chaque ligne du tableau affiche le résultat final de chaque négociation. L'optimum de chaque agent n'est plus pris comme une valeur absolue mais comme une moyenne d'une loi normale dont l'écart-type est spécifié.

5) Pour écraser le document Excell original, revenir au Menu et cliquer sur enregistrer.

6) L'onglet A propos affiche les informations relatives au programme. 
