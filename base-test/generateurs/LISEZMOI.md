# Générateurs des fichiers de test

Lancer depuis `nexus-v/programme`, avec l'environnement Python de NEXUS-V. Les scripts lisent `../base/config.json` sans le modifier.

| Script | Produit | Commande |
| --- | --- | --- |
| `make_modeles.py` | Les modèles chargeables du tome MODELE, avec leur mode d'emploi : exigences, plan, STDHL format 1 et 2, fiches XML (format actuel et format TestSuite) | `python ../../base-test/generateurs/make_modeles.py ../../base-test` |
| `make_stdhl_top.py` | `stdhl/2026-01-01_TEST_STDHL_FCTk_rev1.docx`, à partir des exigences et des plans | `python ../../base-test/generateurs/make_stdhl_top.py ../../base-test ../../base-test/generateurs/gabarits/gabarit-stdhl.docx` |
| `make_fiches_top.py` | `fiches/2026-01-01_TEST_FICHES_FCTk_rev1.xml`, à partir des STDHL | `python ../../base-test/generateurs/make_fiches_top.py ../../base-test` |
| `make_stdhl_format2.py` | `stdhl/2026-11-11-STDHL-FCT1-format2.docx` (STDHL format 2 de FCT1, taux d'incohérence `INCONSISTENCY`) | `python ../../base-test/generateurs/make_stdhl_format2.py ../../base-test ../../base-test/stdhl/2026-11-11-STDHL-FCT1-format2.docx` |

Les fichiers d'exigences (`exigences/2026-01-01-fctk.xlsx`) et les plans (`plans/2026-01-01_PA_FCTk_A-plan.docx`) sont des fichiers de référence. Ils ne sont plus régénérés, car ils ont été retouchés depuis leur génération.

`gabarits/gabarit-stdhl.docx` ne sert que de mise en page pour `make_stdhl_top.py`. Ce n'est pas un modèle de format.

Tous les documents Word générés sont en A4 portrait.
