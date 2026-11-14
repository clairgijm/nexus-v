-- Définition des tables de la base NEXUS-V (nexus-v.db), lue au démarrage par nexus_db.py.
-- La version ne change que lorsqu'une migration est décidée ; une base d'une autre version est refusée.
-- version : 1.0

-- Réglages internes : version de la base, prochain numéro de fichier, référentiels initialisés
CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- Référentiel : tomes auxquels sont rattachés exigences et plans
CREATE TABLE IF NOT EXISTS list_tomes (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE COLLATE NOCASE,
    description TEXT NOT NULL DEFAULT '',     -- Contenu du tome, précisé par l'utilisateur (Configuration > Tomes)
    color TEXT NOT NULL DEFAULT ''            -- couleur dans les camemberts (#rrggbb) ; vide : tirée à la lecture
);

-- Référentiel : plateformes de test
CREATE TABLE IF NOT EXISTS list_plateformes (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE COLLATE NOCASE,
    color TEXT NOT NULL DEFAULT ''            -- couleur dans les camemberts (#rrggbb) ; vide : tirée à la lecture
);

-- Référentiel : codes IADT
CREATE TABLE IF NOT EXISTS list_iadt (
    id INTEGER PRIMARY KEY,
    code TEXT NOT NULL UNIQUE COLLATE NOCASE,
    label TEXT NOT NULL DEFAULT '',
    color TEXT NOT NULL DEFAULT ''            -- couleur dans les camemberts (#rrggbb) ; vide : tirée à la lecture
);

-- Référentiel : codes MOC
CREATE TABLE IF NOT EXISTS list_moc (
    id INTEGER PRIMARY KEY,
    code TEXT NOT NULL UNIQUE COLLATE NOCASE,
    description TEXT NOT NULL DEFAULT '',
    color TEXT NOT NULL DEFAULT ''            -- couleur dans les camemberts (#rrggbb) ; vide : tirée à la lecture
);

-- Tous les fichiers chargés ; source donne leur type
CREATE TABLE IF NOT EXISTS files (
    id INTEGER PRIMARY KEY,
    source TEXT NOT NULL CHECK (source IN ('requirements', 'plan', 'stdhl', 'xml', 'traceability')),
    name TEXT NOT NULL,
    stored_name TEXT NOT NULL,
    tome_id INTEGER REFERENCES list_tomes (id),
    stdhl_name TEXT NOT NULL DEFAULT '',                                                             -- STDHL : nom tiré du nom du fichier
    issues TEXT NOT NULL DEFAULT '',                                                                 -- STDHL : anomalies du document, une par ligne
    platform_id INTEGER REFERENCES list_plateformes (id),                                            -- Fiches d'essai : plateforme d'où vient le fichier
    format TEXT NOT NULL DEFAULT ''                                                                  -- Format reconnu à l'import (STDHL : par tome, unique)
);

-- Valeurs de la colonne STDHL des fichiers d'exigences
CREATE TABLE IF NOT EXISTS list_stdhl (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE COLLATE NOCASE  -- Valeur trouvée dans la colonne STDHL
);

-- Exigences
CREATE TABLE IF NOT EXISTS requirements (
    id INTEGER PRIMARY KEY,
    req_id TEXT NOT NULL,                                    -- Numéro de l'exigence
    file_id INTEGER REFERENCES files (id) ON DELETE CASCADE,
    stdhl_id INTEGER REFERENCES list_stdhl (id),
    tome_id INTEGER NOT NULL REFERENCES list_tomes (id),
    title TEXT NOT NULL DEFAULT '',
    iadt_id INTEGER REFERENCES list_iadt (id),               -- Méthode de vérification (référence)
    moc_id INTEGER REFERENCES list_moc (id),                 -- Moyen de vérification (référence)
    delegation_id INTEGER REFERENCES list_plateformes (id),  -- Plateforme qui vérifie l'exigence (référence)
    validation TEXT NOT NULL DEFAULT 'out_of_scope'          -- Validation dans la release de la base (« Validation releases ») :
        CHECK (validation IN ('to_qualify', 'qualified', 'out_of_scope')),  -- à qualifier, qualifiée, hors scope
    manual TEXT CHECK (manual IN ('added', 'modified')),     -- Modifiée à la main depuis le dernier import
    deleted INTEGER NOT NULL DEFAULT 0,                      -- Supprimée à la main : affichée barrée, non vérifiée
    UNIQUE (req_id)
);
CREATE INDEX IF NOT EXISTS requirements_file ON requirements (file_id);

-- Plans d'acceptation, une ligne par exigence du tome du plan (et par code IADT et MOC)
CREATE TABLE IF NOT EXISTS plan_acceptation (
    id INTEGER PRIMARY KEY,
    requirement_id INTEGER NOT NULL REFERENCES requirements (id) ON DELETE CASCADE,
    iadt_id INTEGER REFERENCES list_iadt (id),                                           -- Cellule vide : pas de valeur
    moc_id INTEGER REFERENCES list_moc (id),
    delegation_id INTEGER REFERENCES list_plateformes (id),                                -- Cellule vide : délégation par défaut
    file_id INTEGER NOT NULL REFERENCES files (id) ON DELETE CASCADE,
    manual TEXT,                                -- added, modified (à la main ou corrigée à l'import), deleted ; vide : lue dans le document
    deleted INTEGER NOT NULL DEFAULT 0          -- 1 : supprimée à la main ou ignorée à l'import (barrée, sans effet)
);


-- Couverture des exigences par les STDHL : une ligne par exigence citée dans un STDHL (numéros absents de la base
-- non enregistrés : anomalie du document)
CREATE TABLE IF NOT EXISTS stdhl_exi (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files (id) ON DELETE CASCADE,
    requirement_id INTEGER NOT NULL REFERENCES requirements (id) ON DELETE CASCADE,
    test_id INTEGER REFERENCES fe_tests (id) ON DELETE CASCADE,          -- Test (fiche SYS_) demandé ; vide hors fiche
    deleted INTEGER NOT NULL DEFAULT 0,   -- 1 : supprimée par l'utilisateur dans la page STDHL (barrée, non vérifiée)
    UNIQUE (file_id, requirement_id, test_id)
);

-- Référentiel des tests (fiches SYS_), communs aux STDHL (tests demandés) et aux fiches d'essai (tests réalisés) ;
-- le test sans nom porte les fichiers Excel des autres plateformes. Un test qui n'est plus cité est supprimé.
CREATE TABLE IF NOT EXISTS fe_tests (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE                                             -- Nom du test
);

-- Test cases des tests : ceux des fiches d'essai XML
CREATE TABLE IF NOT EXISTS fe_test_case (
    id INTEGER PRIMARY KEY,
    test_id INTEGER NOT NULL REFERENCES fe_tests (id) ON DELETE CASCADE,
    test_case TEXT NOT NULL DEFAULT '',                                   -- Nom du test case
    UNIQUE (test_id, test_case)
);

-- Test cases demandés par un STDHL (format unique) : une fiche d'essai dont le test case n'y est pas est ignorée
CREATE TABLE IF NOT EXISTS stdhl_test_case (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files (id) ON DELETE CASCADE,
    test_case_id INTEGER NOT NULL REFERENCES fe_test_case (id) ON DELETE CASCADE,
    UNIQUE (file_id, test_case_id)
);

-- Couverture des exigences par les Fiches d'essai : une ligne par exigence, rattachée toujours à un test et, pour les
-- fiches d'essai XML, aussi à un test case dont elle prend le résultat ; numéros absents de la base non enregistrés
CREATE TABLE IF NOT EXISTS fiches_essai (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files (id) ON DELETE CASCADE,
    requirement_id INTEGER NOT NULL REFERENCES requirements (id) ON DELETE CASCADE,
    test_id INTEGER NOT NULL REFERENCES fe_tests (id) ON DELETE CASCADE,  -- Test
    test_case_id INTEGER REFERENCES fe_test_case (id) ON DELETE CASCADE,  -- Test case (fiches d'essai XML), facultatif
    result TEXT NOT NULL DEFAULT 'NOT_RUN' CHECK (result IN ('NOT_RUN', 'PASSED', 'FAILED'))  -- Résultat du test case
);
CREATE UNIQUE INDEX IF NOT EXISTS fiches_essai_test ON fiches_essai (file_id, requirement_id, test_id)
    WHERE test_case_id IS NULL;
CREATE UNIQUE INDEX IF NOT EXISTS fiches_essai_test_case ON fiches_essai (file_id, requirement_id, test_case_id)
    WHERE test_case_id IS NOT NULL;

-- Un test case que plus aucune fiche ne cite est retiré ; un test que plus aucun STDHL, aucune fiche ni aucun test
-- case ne cite est retiré du référentiel.
CREATE TRIGGER IF NOT EXISTS stdhl_exi_test_unused AFTER DELETE ON stdhl_exi WHEN OLD.test_id IS NOT NULL
BEGIN
    DELETE FROM fe_tests WHERE id = OLD.test_id
        AND NOT EXISTS (SELECT 1 FROM stdhl_exi WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fiches_essai WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fe_test_case WHERE test_id = OLD.test_id);
END;
CREATE TRIGGER IF NOT EXISTS fiches_essai_test_unused AFTER DELETE ON fiches_essai
BEGIN
    DELETE FROM fe_tests WHERE id = OLD.test_id
        AND NOT EXISTS (SELECT 1 FROM stdhl_exi WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fiches_essai WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fe_test_case WHERE test_id = OLD.test_id);
END;
CREATE TRIGGER IF NOT EXISTS fiches_essai_test_case_unused AFTER DELETE ON fiches_essai WHEN OLD.test_case_id IS NOT NULL
BEGIN
    DELETE FROM fe_test_case WHERE id = OLD.test_case_id
        AND NOT EXISTS (SELECT 1 FROM fiches_essai WHERE test_case_id = OLD.test_case_id)
        AND NOT EXISTS (SELECT 1 FROM stdhl_test_case WHERE test_case_id = OLD.test_case_id);
END;
CREATE TRIGGER IF NOT EXISTS stdhl_test_case_unused AFTER DELETE ON stdhl_test_case
BEGIN
    DELETE FROM fe_test_case WHERE id = OLD.test_case_id
        AND NOT EXISTS (SELECT 1 FROM fiches_essai WHERE test_case_id = OLD.test_case_id)
        AND NOT EXISTS (SELECT 1 FROM stdhl_test_case WHERE test_case_id = OLD.test_case_id);
END;
CREATE TRIGGER IF NOT EXISTS fe_test_case_test_unused AFTER DELETE ON fe_test_case
BEGIN
    DELETE FROM fe_tests WHERE id = OLD.test_id
        AND NOT EXISTS (SELECT 1 FROM stdhl_exi WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fiches_essai WHERE test_id = OLD.test_id)
        AND NOT EXISTS (SELECT 1 FROM fe_test_case WHERE test_id = OLD.test_id);
END;
