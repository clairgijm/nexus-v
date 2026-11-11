"""Loadable templates of every file type read by NEXUS-V, for the tome MODELE (release A).

Each template starts with a « Mode d'emploi » explaining the expected format, then holds a complete example that can be
imported as is: the requirements of the tome MODELE, its acceptance plan, its STDHL (format 1 and format 2) and its
Fiches d'Essai (current XML format and XML format TestSuite). The explanations only use placeholders such as
SSS_<TOME>_REQ_<n°>, which the reading ignores: only the example numbers are imported.
The Word templates are built on the styles of 2026-11-11-STDHL-template2.docx.
Usage: python make_modeles.py <base-test>   (writes into exigences/, plans/, stdhl/ (format 1 in stdhl/perTome/) and fiches/)
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(sys.argv[1])
SKELETON = BASE / "stdhl" / "2026-11-11-STDHL-template2.docx"

# The example: requirements of the tome MODELE, their plan row and the test that covers them.
#   number, title, releases, STDHL, IADT, MOC, delegation ("" = TOP, default platform), test, test case
REQS = [
    ("SSS_MODELE_REQ_001", "Le système horodate chaque événement", "A", "MODELE", "D", "MOC4", "",
     "SYS_PERF_MODELE_001", 1),
    ("SSS_MODELE_REQ_002", "Le système transmet les alarmes à la plateforme", "A", "", "T", "MOC4", "PLT1_N-1", "", 0),
    ("SSS_MODELE_REQ_003", "Le temps de réponse est inférieur à 2 s", "A", "MODELE", "A", "MOC2, MOC3", "",
     "SYS_PERF_MODELE_001", 2),
    ("SSS_MODELE_REQ_004", "L'IHM affiche l'état des équipements", "A", "MODELE", "I", "MOC7", "",
     "SYS_COM_MODELE_002", 1),
    ("SSS_MODELE_REQ_005", "Exigence d'information, sans essai", "A", "", "N/A", "N/A", "", "", 0),
    ("SSS_MODELE_REQ_006", "Les messages sont acquittés par l'opérateur", "A", "MODELE", "T", "MOC4", "",
     "SYS_COM_MODELE_002", 1),
    ("SSS_MODELE_REQ_007", "La liaison est rétablie après coupure", "A, B", "MODELE", "T", "MOC4", "",
     "SYS_COM_MODELE_002", 2),
    ("HQ1/2_WF_001", "Le flux de travail est journalisé", "A", "MODELE", "T", "MOC4", "", "SYS_COM_MODELE_002", 2),
]
TESTS = {"SYS_PERF_MODELE_001": "Performances", "SYS_COM_MODELE_002": "Communications"}
FORMAT2 = {"SYS_PERF_MODELE_001": "PERFORMANCES_001", "SYS_COM_MODELE_002": "COMMUNICATIONS_002"}


def covered(test, case=None):
    """Requirements of a test (and test case) in the example."""
    return [r for r in REQS if r[7] == test and (case is None or r[8] == case)]


# ---------------------------------------------------------------- Word helpers
def run(text, bold=False, italic=False):
    props = ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "")
    return f'<w:r>{f"<w:rPr>{props}</w:rPr>" if props else ""}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


# The style Titre3 of the template indents its paragraphs by 19.5 cm and moves the text after the number to a tab at
# 21.4 cm, off the page: the title is in the table of contents but cannot be seen. Each title cancels both itself.
TITLE3_FIX = ('<w:tabs><w:tab w:val="clear" w:pos="12137"/><w:tab w:val="num" w:pos="1080"/></w:tabs>'
              '<w:ind w:left="0" w:firstLine="0"/>')


def p(text="", style="", bold=False, italic=False):
    """A paragraph; `text` may be a list of (text, bold) runs."""
    fix = TITLE3_FIX if style == "Titre3" else ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>{fix}</w:pPr>' if style else ""
    runs = "".join(run(t, b, italic) for t, b in text) if isinstance(text, list) else run(text, bold, italic)
    return f"<w:p>{ppr}{runs}</w:p>"


def rule(label, text, style="Paragraphedeliste"):
    """An explanation line « label : text », the label in bold."""
    return p([(label + " : ", True), (text, False)], style)


def table(rows, header=True, spans=()):
    """A Word table with borders; `spans`: row indexes made of one cell spanning every column."""
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    borders = "".join(border.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    width = max(len(r) for r in rows)
    out = []
    for k, row in enumerate(rows):
        if k in spans:
            cells = f'<w:tc><w:tcPr><w:gridSpan w:val="{width}"/></w:tcPr>{p(row[0], bold=True)}</w:tc>'
        else:
            cells = "".join(f"<w:tc>{p(c, bold=header and k == 0)}</w:tc>" for c in row)
        out.append(f"<w:tr>{cells}</w:tr>")
    grid = "".join('<w:gridCol w:w="2000"/>' for _ in range(width))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            f"<w:tblGrid>{grid}</w:tblGrid>" + "".join(out) + "</w:tbl>")


def a4_portrait(xml):
    """Every section of the document in A4 portrait (the skeleton ends with a landscape section)."""
    return re.sub(r"<w:pgSz [^>]*/>", '<w:pgSz w:w="11906" w:h="16838" w:code="9"/>', xml)


def write_docx(path, body):
    """A .docx with the styles of the skeleton and this body."""
    with zipfile.ZipFile(SKELETON) as z, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "word/document.xml":
                x = data.decode("utf-8")
                a, b = x.index("<w:body>") + len("<w:body>"), x.rindex("<w:sectPr")
                data = a4_portrait(x[:a] + "".join(body) + x[b:]).encode("utf-8")
            w.writestr(info, data)
    print("écrit :", path.relative_to(BASE))


def notice(title, lines):
    """The « Mode d'emploi » section of a Word template."""
    return ([p(title, "Titre"), p("Mode d'emploi", "Titre1"),
             p("Ce modèle peut être chargé tel quel dans NEXUS-V : il décrit le format attendu, puis en donne un exemple "
               "complet pour le tome MODELE, en release A. Les textes en majuscules entre < > sont à remplacer ; ils ne "
               "sont pas lus comme des numéros d'exigence. Les libellés, en-têtes et expressions régulières cités sont "
               "ceux de la configuration par défaut (onglet Configuration de NEXUS-V).", italic=True)]
            + [rule(label, text) for label, text in lines])


# ---------------------------------------------------------------- requirements (.xlsx)
def requirements():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exigences"
    ws.append(["ID", "Titre", "Release", "STDHL"])
    for r in REQS:
        ws.append([r[0], r[1], r[2], r[3]])
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
    for col, w in zip("ABCD", (24, 52, 12, 14)):
        ws.column_dimensions[col].width = w
    help_ = wb.create_sheet("Mode d'emploi")
    lines = [
        ("Modèle de fichier d'exigences NEXUS-V", ""),
        ("", "Ce classeur peut être importé tel quel : il crée le tome MODELE. NEXUS-V lit la première feuille ; "
             "cette feuille-ci n'est pas lue."),
        ("Nom du fichier", "<AAAA-MM-JJ>-<TOME>.xlsx, par exemple 2026-01-01-MODELE.xlsx : le tome est le texte après "
                           "la date, jusqu'au tiret, au souligné ou à l'extension suivant (requirements_tome_regex)."),
        ("Un seul tome", "un fichier d'exigences porte un seul tome (une seule fonction)."),
        ("Ligne 1", "les en-têtes, dans n'importe quel ordre : ID, Titre, Release, STDHL (headers_requirements). "
                    "Un en-tête absent fait refuser le fichier ; les autres colonnes sont ignorées."),
        ("ID", "numéro de l'exigence, au format SSS_<TOME>_REQ_<n°> (avec ou sans T<n>_ après SSS_), ou "
               "HQ1/2_WF_<n°> suivi ou non de _REQ_<n°> (requirement_regex). Un numéro présent sur plusieurs lignes "
               "est signalé à l'import : l'utilisateur choisit la ligne gardée."),
        ("Titre", "texte libre."),
        ("Release", "une ou plusieurs releases séparées par une virgule (release_separator_regex), par exemple « A, B ». "
                    "À l'import, l'utilisateur choisit les releases gardées ; une exigence hors de ces releases peut "
                    "être forcée ou ignorée."),
        ("STDHL", "nom du STDHL qui spécifie l'exigence (par défaut, le tome). Vide quand l'exigence n'est pas "
                  "spécifiée dans un STDHL de la plateforme TOP. Plusieurs valeurs : séparées par , ; / ou un retour "
                  "à la ligne (stdhl_separator_regex). Vérifié seulement si « Utiliser la colonne STDHL » est coché "
                  "dans Configuration > STDHL."),
        ("Exemple", "8 exigences : 001 à 007 et HQ1/2_WF_001. 002 est déléguée à la plateforme PLT1_N-1 et 005 est "
                    "N/A dans le plan : ni l'une ni l'autre n'a de STDHL."),
    ]
    for label, text in lines:
        help_.append([label, text])
    help_["A1"].font = Font(bold=True, size=14)
    for row in help_.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    help_.column_dimensions["A"].width, help_.column_dimensions["B"].width = 18, 100
    for sheet in wb.worksheets:   # printed in A4 portrait
        sheet.page_setup.paperSize, sheet.page_setup.orientation = sheet.PAPERSIZE_A4, "portrait"
        sheet.page_setup.fitToWidth, sheet.sheet_properties.pageSetUpPr.fitToPage = 1, True
        sheet.page_setup.fitToHeight = 0
    path = BASE / "exigences" / "2026-01-01-MODELE.xlsx"
    wb.save(path)
    print("écrit :", path.relative_to(BASE))


# ---------------------------------------------------------------- acceptance plan (.docx)
def plan():
    body = notice("Plan d'acceptation – modèle NEXUS-V", [
        ("Nom du fichier", "<AAAA-MM-JJ>_PA_<TOME>_<release>-<texte>.docx, par exemple "
                           "2026-01-01_PA_MODELE_A-modele.docx : le tome est le texte entre « _PA_ » et le "
                           "souligné suivant (plan_tome_regex)."),
        ("Tableaux lus", "tout tableau dont la première ligne contient les en-têtes Exigence, IADT, MOC et Délégation "
                         "(headers_plan) ; la colonne Méthode est lue mais non enregistrée. Les autres tableaux du "
                         "document sont ignorés ; un tableau auquel il manque un de ces en-têtes est signalé."),
        ("Exigence", "le numéro de l'exigence, suivi ou non de son titre : SSS_<TOME>_REQ_<n°> <titre>. Une exigence "
                     "inconnue du tome, ou hors de la release du plan, est à ignorer au chargement."),
        ("IADT", "I (inspection), A (analyse), D (démonstration), T (test) ou N/A ; plusieurs codes séparés par , ; + "
                 "ou un espace (code_separator_regex)."),
        ("MOC", "le moyen de conformité attendu pour l'IADT (règles IADT / MOC de la configuration) : D et T → MOC4, "
                "I → MOC7, A → MOC2 ou MOC3, N/A → N/A. Une valeur hors règle est à corriger au chargement."),
        ("Délégation", "la plateforme qui teste l'exigence ; vide = la plateforme par défaut (TOP). Une plateforme "
                       "inconnue de la configuration, ou plusieurs délégations, sont à régler au chargement."),
        ("N/A", "une exigence dont toutes les lignes sont IADT N/A et MOC N/A n'est pas testée : elle n'est attendue "
                "ni dans les STDHL ni dans les fiches."),
        ("Rattachement", "un plan porte sur un seul tome ; un tableau par thème (feature), sous un titre, est l'usage."),
    ])
    header = ["Exigence", "Méthodes de test V&V", "IADT", "MOC", "Délégation"]
    body += [p("Exemple : tome MODELE, release A", "Titre1")]
    for theme, numbers in (("Fonctions de base", REQS[:4]), ("Exploitation", REQS[4:])):
        rows = [header] + [[f"{r[0]} {r[1]}", f"V&V [{r[6] or 'TOP'}] : méthode de test et moyens", r[4], r[5], r[6]]
                           for r in numbers]
        body += [p(theme, "Titre2"), table(rows)]
    write_docx(BASE / "plans" / "2026-01-01_PA_MODELE_A-modele.docx", body)


# ---------------------------------------------------------------- STDHL format 1 (.docx)
def stdhl_format1():
    body = notice("STDHL (format 1, en tableaux) – modèle NEXUS-V", [
        ("Nom du fichier", "<AAAA-MM-JJ>_<texte>_STDHL_<nom du STDHL>_rev<texte>.docx, par exemple "
                           "2026-01-01_STDHL_MODELE_rev1.docx : le nom du STDHL est le texte entre « STDHL_ » et "
                           "« _rev » (stdhl_name_regex). Un nom hors format est signalé, sans bloquer."),
        ("Rattachement", "au chargement, le STDHL est rattaché à un tome, ou à aucun : il vaut alors pour tous les "
                         "tomes. Un STDHL de même nom remplace le précédent."),
        ("Test", "un paragraphe qui commence par le nom du test : SYS_<TYPE>_<TOME>_<n°>, éventuellement précédé d'un "
                 "numéro de titre (fiche_regex). Tous les tableaux qui suivent appartiennent à ce test, jusqu'au "
                 "test suivant. Le nom est comparé, en entier, à celui des fiches d'essai."),
        ("Test case", "un paragraphe « TEST CASE #<n> : <titre> » avant ses tableaux (non lu, pour le lecteur)."),
        ("Tableaux lus", "ceux dont la première ligne contient les en-têtes Description du test, Résultat attendu et "
                         "Exigences associées (headers_stdhl) ; les autres tableaux sont ignorés. Une ligne de phase "
                         "(cellule fusionnée « Phase <n> ») est permise."),
        ("Exigences associées", "les numéros des exigences vérifiées par la ligne, séparés par un espace, une virgule "
                                "ou un retour à la ligne. Un numéro absent de la base, ou une exigence N/A dans les "
                                "plans, n'est pas enregistré et est signalé."),
        ("Format", "un document sans aucun tableau valide est lu comme un STDHL de format 2 (voir le modèle "
                   "2026-01-01_STDHL_MODELE-F2_rev1.docx)."),
    ])
    body.append(p("Description des tests", "Titre1"))
    header = ["Description du test", "Résultat attendu", "Exigences associées"]
    for k, (test, theme) in enumerate(TESTS.items(), 1):
        body += [p(f"{test}", "Titre2"), p(f"Objectif : vérification des {theme.lower()} de MODELE"),
                 p("Procédure de test", "Titre3")]
        for case in (1, 2):
            reqs = covered(test, case)
            rows = [header, ["Phase 1"], ["Préparer le banc de test", "Banc prêt", ""]]
            rows += [[f"Vérifier : {r[1].lower()}", "Conforme à l'exigence", r[0]] for r in reqs]
            body += [p(f"TEST CASE #{case} : {theme.lower()}, partie {case}"), table(rows, spans=(1,))]
    write_docx(BASE / "stdhl" / "perTome" / "2026-01-01_STDHL_MODELE_rev1.docx", body)


# ---------------------------------------------------------------- STDHL format 2 (.docx)
def stdhl_format2():
    body = notice("STDHL (format 2, en texte) – modèle NEXUS-V", [
        ("Format en préparation", "ce format n'est pas encore lu par NEXUS-V : ce modèle fixe le format attendu."),
        ("Usage", "un seul document regroupe tous les STDHL ; il est lu quand aucun de ses tableaux n'a les en-têtes "
                  "du format 1. Il peut contenir d'autres tableaux (historique, moyens d'essai), ignorés."),
        ("Nom du fichier", "<AAAA-MM-JJ>_<texte>_STDHL_<nom du STDHL>_rev<texte>.docx, comme le format 1."),
        ("Test", "un titre « <TYPE DE TEST>_<n°> : <titre du test> » ; le type décrit l'essai effectué (30 caractères "
                 "au plus), le numéro est unique dans le document. Le nom du test est le texte avant le premier « : », "
                 "comparé en entier à celui des fiches d'essai. La table des matières, qui répète ces titres, est "
                 "ignorée."),
        ("Applicabilité", "« Applicabilité : RELEASE <release> » ; une release absente, ou différente de celle de la "
                          "base, sera signalée sans bloquer le chargement."),
        ("Exigences associées", "le libellé « Exigences associées : », puis une exigence par paragraphe : "
                                "SSS_<TOME>_REQ_<n°> : <titre>. La liste se termine à « Conditions et configuration "
                                "initiales : » ou au test suivant."),
        ("Test cases", "après « DESCRIPTION DE LA FICHE DE TEST », chaque test case commence par « TEST CASE#<n> : "
                       "<titre> », suivi de ses phases « Phase <n> : <action> : <exigences citées> ». Les exigences sont "
                       "gérées au test case : une exigence associée citée dans aucun test case, ou une exigence d'un "
                       "test case absente des exigences associées, seront signalées."),
    ])
    body += [p("Historique des versions", "Titre1"),
             table([["Version", "Date", "Auteur", "Modification"], ["1.0", "2026-01-01", "V&V", "Création du modèle"]]),
             p("Description des tests", "Titre1"), p("Tests de la fonction MODELE", "Titre2")]
    for test, theme in TESTS.items():
        reqs = covered(test)
        body += [p(f"{FORMAT2[test]} : {theme.lower()} de MODELE", "Titre3"), p("État : Revue"),
                 p(f"Objectif : vérifier les {theme.lower()} de MODELE"), p("Applicabilité : RELEASE A"),
                 p("Exigences associées :")]
        body += [p(f"{r[0]} : {r[1]}", "Paragraphedeliste") for r in reqs]
        body += [p("Conditions et configuration initiales :"),
                 p("Banc de test en configuration nominale", "Paragraphedeliste"),
                 p("Moyens d'essai :"), table([["Moyen", "Configuration"], ["Banc de test V&V", "Nominale"]]),
                 p("DESCRIPTION DE LA FICHE DE TEST")]
        for case in (1, 2):
            body.append(p(f"TEST CASE#{case} : {theme.lower()}, partie {case}"))
            body.append(p("Phase 1 : préparer le banc de test", "Paragraphedeliste"))
            for k, r in enumerate(covered(test, case), 2):
                body.append(p(f"Phase {k} : vérifier que {r[1][0].lower() + r[1][1:]} : {r[0]}", "Paragraphedeliste"))
    write_docx(BASE / "stdhl" / "2026-01-01_STDHL_MODELE-F2_rev1.docx", body)


# ---------------------------------------------------------------- Fiches d'Essai (.xml)
def fiches_format1():
    lines = ['<?xml version="1.0" encoding="UTF-8"?>',
             "<!-- Fiches d'essai – modèle NEXUS-V (format XML actuel). Ce fichier peut être chargé tel quel.",
             "",
             "  Nom du fichier : libre, par exemple <AAAA-MM-JJ>_TEST_FICHES_<TOME>_rev<n>.xml. Au chargement, le",
             "    fichier est rattaché à un tome, ou à aucun (il vaut alors pour tous les tomes).",
             "  Balises (configurables, Configuration > Fiches d'essai) :",
             "    <test>          un test ; son texte est le nom du test, SYS_<TYPE>_<TOME>_<n°>, comparé en entier",
             "                    au nom du test demandé par le STDHL ;",
             "    <test_case>     un test case du test ; son texte est le titre du test case ;",
             "    <requirements>  les exigences vérifiées par le test case, séparées par une virgule (une ou",
             "                    plusieurs balises par test case).",
             "  Numéros : SSS_<TOME>_REQ_<n°> ou HQ1/2_WF_<n°> (requirement_regex). Un test case sans exigence",
             "  n'est pas lu. La racine et les autres balises sont libres.",
             "  Exemple : les deux tests du STDHL MODELE, avec les mêmes exigences. -->",
             "<fiches>"]
    for test, theme in TESTS.items():
        lines.append(f"  <test>{test}")
        for case in (1, 2):
            reqs = ", ".join(r[0] for r in covered(test, case))
            lines += [f"    <test_case>TEST CASE #{case} : {theme.lower()}, partie {case}",
                      f"      <requirements>{escape(reqs)}</requirements>", "    </test_case>"]
        lines.append("  </test>")
    lines.append("</fiches>")
    path = BASE / "fiches" / "2026-01-01_TEST_FICHES_MODELE_rev1.xml"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("écrit :", path.relative_to(BASE))


def fiches_format2():
    def props(**values):
        out = ["<Properties>"]
        for name, value in values.items():
            out.append(f'<Property name="{name}">{escape(value)}</Property>' if value else f'<Property name="{name}"/>')
        return "".join(out) + "</Properties>"
    lines = ['<?xml version="1.0" encoding="UTF-8"?>',
             "<!-- Fiches d'essai – modèle NEXUS-V (format XML TestSuite, exporté par l'outil de campagne).",
             "",
             "  Format en préparation : ce format n'est pas encore lu par NEXUS-V ; ce modèle fixe ce qui sera lu.",
             "  Structure : Project / Tcr / Campaign / TestSuites / TestSuite / TestCases / TestCase / TestSteps /",
             "    TestStep. La partie Configuration du Project n'est pas lue.",
             "  TestSuite : un test ; son attribut name (et sa propriété Title) est le nom du test,",
             "    SYS_<TYPE>_<TOME>_<n°>, comparé en entier au nom du test demandé par le STDHL.",
             "  TestCase : un test case ; sa propriété Title est son titre.",
             "  Exigences : la propriété VerifiedRequirements, au niveau du TestSuite, du TestCase ou du TestStep ;",
             "    plusieurs numéros séparés par une virgule. Les exigences d'un TestStep appartiennent à son TestCase.",
             "  Campaign : l'attribut name porte la release (« Roadmap A »).",
             "  Exemple : les deux tests du STDHL MODELE, exigences portées par les TestStep. -->",
             '<Project name="MODELE" path="">', "  <Configuration/>", "  <Tcr>", '    <Campaign name="Roadmap A">',
             "      <TestSuites>"]
    uid = 0
    for test, theme in TESTS.items():
        lines += [f'        <TestSuite name="{test}">',
                  "          " + props(Title=test, DefinitionVersion="1", RunDate="2026-01-01", ManualOrAuto="MANUAL",
                                       Method="T", VerifiedRequirements=""),
                  "          <TestCases>"]
        for case in (1, 2):
            uid += 1
            lines += [f'            <TestCase uid="{uid}">',
                      "              " + props(Title=f"{theme.lower()}, partie {case}", RunDate="2026-01-01",
                                               RunningStatus="RUN", VerifiedRequirements=""),
                      "              <TestSteps>"]
            uid += 1
            lines += [f'                <TestStep uid="{uid}">' + props(Description="Préparer le banc de test",
                                                                        VerifiedRequirements="", StepType="CONTEXT")
                      + "</TestStep>"]
            for r in covered(test, case):
                uid += 1
                lines += [f'                <TestStep uid="{uid}">' + props(Description=f"Vérifier : {r[1].lower()}",
                                                                            VerifiedRequirements=r[0], StepType="ACTION")
                          + "</TestStep>"]
            lines += ["              </TestSteps>", "            </TestCase>"]
        lines += ["          </TestCases>", "        </TestSuite>"]
    lines += ["      </TestSuites>", "    </Campaign>", "  </Tcr>", "</Project>"]
    path = BASE / "fiches" / "2026-01-01_FE_MODELE_format2.xml"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("écrit :", path.relative_to(BASE))


requirements()
plan()
stdhl_format1()
stdhl_format2()
fiches_format1()
fiches_format2()
