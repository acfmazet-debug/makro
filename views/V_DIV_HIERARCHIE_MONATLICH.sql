-- =============================================================================
-- View : PROD_SDP_UL_DAEF_REPORTING.DAEV.V_DIV_HIERARCHIE_MONATLICH
-- Quelle: PROD_SDP_AL_PERS.AL_DIV_SHARED.DIV_TBL_HIERARCHIE
--
-- Zweck:
--   Eine Zeile pro Vermittler (BO) pro Jahr (JAHR) und Monat (MONAT).
--   Da die Quelltabelle tägliche Einträge liefert, wird per "Mehrheitsentscheid"
--   (Modal-Wert) die häufigste Zuordnungskombination des Monats gewählt.
--   Damit werden einzelne Fehltage / Fehlzuordnungen automatisch herausgefiltert.
--
-- Filter:
--   - VTWS IN (11017, 11028, 11025, 11020, 11021, 11018, 11019)
--   - JAHR > '2016'
--
-- Logik:
--   1. Alle Tageszeilen filtern (WHERE)
--   2. Alle inhaltlichen Spalten gruppieren + Anzahl Tage zählen
--   3. Pro BO/JAHR/MONAT die Gruppe mit den meisten Tagen behalten (QUALIFY)
--
-- Spaltenbehandlung:
--   - STICHTAG         : ANY_VALUE  (Tagesschlüssel – nicht sinnvoll aggregierbar)
--   - LSALES/LPERS/LCOR: ANY_VALUE  (bis 16 MB, nicht sinnvoll gruppierbar)
--   - Alle anderen     : explizit im GROUP BY
-- =============================================================================

CREATE OR REPLACE VIEW PROD_SDP_UL_DAEF_REPORTING.DAEV.V_DIV_HIERARCHIE_MONATLICH AS

WITH gezaehlt AS (
    SELECT
        -- Zeitdimension
        JAHR,
        MONAT,

        -- Vermittler (Basis)
        BO,
        BO_NAME,
        BO_VVH,
        BO_SEIT,
        PA_ANTEIL,

        -- Zugeordnete Konten (ZKTO1-8)
        ZKTO1,        ZKTO1_NAME,   ZKTO1_VVH,
        ZKTO2,        ZKTO2_NAME,   ZKTO2_VVH,
        ZKTO3,        ZKTO3_NAME,   ZKTO3_VVH,
        ZKTO4,        ZKTO4_NAME,   ZKTO4_VVH,
        ZKTO5,        ZKTO5_NAME,   ZKTO5_VVH,
        ZKTO6,        ZKTO6_NAME,   ZKTO6_VVH,
        ZKTO7,        ZKTO7_NAME,   ZKTO7_VVH,
        ZKTO8,        ZKTO8_NAME,   ZKTO8_VVH,

        -- Eingangskontakt
        EKTO,         EKTO_VVH,     EKTO_NAME,   EKTO_SEIT,   EKTO_ALTVERS,

        -- Sonstige Zuordnungen
        INSP,
        QUER,
        MKTO,         MKTO_NAME,    MKTO_VVH,
        MSB,          MSB_NAME,
        MAKLCAT,
        MAKSEG,       MAKSEG_TXT,
        VERBAND,      MITGLNUMMER,

        -- IBN / IBSN Hierarchie
        IBN,          IBN_NAME,
        IBSN,         IBSN_NAME,
        FUEHR2_IBN,   FUEHR2_IBN_NAME,
        FUEHR1_IBN,   FUEHR1_IBN_NAME,
        FUEHR2_IBSN,  FUEHR2_IBSN_NAME,
        FUEHR1_IBSN,  FUEHR1_IBSN_NAME,
        VTWEG_IBN,
        VTWEG_IBSN,

        -- MBV Hierarchie
        MBV,              MBV_NAME,         MBV_VVH,
        FUEHR2_MBV,       FUEHR2_MBV_NAME,
        FUEHR1_MBV,       FUEHR1_MBV_NAME,
        VTWEG_MBV,

        -- MBKV Hierarchie
        MBKV,             MBKV_NAME,
        FUEHR2_MBKV,      FUEHR2_MBKV_NAME,
        FUEHR1_MBKV,      FUEHR1_MBKV_NAME,
        VTWEG_MBKV,

        -- OEDSP Hierarchie
        OEDSP,            OEDSP_NAME,
        FUEHR2_OEDSP,     FUEHR2_OEDSP_NAME,
        FUEHR1_OEDSP,     FUEHR1_OEDSP_NAME,
        VTWEG_OEDSP,

        -- MBCEB Hierarchie
        MBCEB,            MBCEB_NAME,
        FUEHR2_MBCEB,     FUEHR2_MBCEB_NAME,
        FUEHR1_MBCEB,     FUEHR1_MBCEB_NAME,
        VTWEG_MBCEB,

        -- MBKI Hierarchie
        MBKI,             MBKI_NAME,
        FUEHR2_MBKI,      FUEHR2_MBKI_NAME,
        FUEHR1_MBKI,      FUEHR1_MBKI_NAME,
        VTWEG_MBKI,

        -- KAMKV Hierarchie
        KAMKV,            KAMKV_NAME,
        FUEHR2_KAMKV,     FUEHR2_KAMKV_NAME,
        FUEHR1_KAMKV,     FUEHR1_KAMKV_NAME,
        VTWEG_KAMKV,

        -- KAMV Hierarchie
        KAMV,             KAMV_NAME,
        FUEHR2_KAMV,      FUEHR2_KAMV_NAME,
        FUEHR1_KAMV,      FUEHR1_KAMV_NAME,
        VTWEG_KAMV,

        -- MBK Hierarchie
        MBK,              MBK_NAME,         MBK_VVH,
        FUEHR2_MBK,       FUEHR2_MBK_NAME,
        FUEHR1_MBK,       FUEHR1_MBK_NAME,
        VTWEG_MBK,

        -- MRK Hierarchie
        MRK,              MRK_NAME,
        FUEHR2_MRK,       FUEHR2_MRK_NAME,
        FUEHR1_MRK,       FUEHR1_MRK_NAME,
        VTWEG_MRK,

        -- MBPAV Hierarchie
        MBPAV,            MBPAV_NAME,
        FUEHR2_MBPAV,     FUEHR2_MBPAV_NAME,
        FUEHR1_MBPAV,     FUEHR1_MBPAV_NAME,
        VTWEG_MBPAV,

        -- Betriebsleiter / Führungskräfte
        BETRLI,           BETRLI_NAME,      BETRLI_VVH,
        FUEHR2,           FUEHR2_NAME,      FUEHR2_TEXT,
        FUEHR1,           FUEHR1_NAME,      FUEHR1_TEXT,

        -- Vertriebsweg / Verkaufsleitung
        VTWS,
        VKL,              VKL_NAME,         VKL_VVH,

        -- Weitere Rollen
        AM,               AM_NAME,
        FBP,              FBP_NAME,
        FBCOR,            FBCOR_NAME,
        MAB,              MAB_NAME,
        RGB,              RGB_NAME,
        FBOED,            FBOED_NAME,
        FBCEB,            FBCEB_NAME,
        LMAB,             LMAB_NAME,

        -- Tagesschlüssel: nicht aggregierbar, ANY_VALUE liefert einen Referenztag
        ANY_VALUE(STICHTAG)    AS STICHTAG,

        -- Listenspalten (bis 16 MB, z.B. JSON-Arrays) – nicht sinnvoll gruppierbar
        ANY_VALUE(LSALES)      AS LSALES,
        ANY_VALUE(LSALES_NAME) AS LSALES_NAME,
        ANY_VALUE(LPERS)       AS LPERS,
        ANY_VALUE(LPERS_NAME)  AS LPERS_NAME,
        ANY_VALUE(LCOR)        AS LCOR,
        ANY_VALUE(LCOR_NAME)   AS LCOR_NAME,

        -- Anzahl Tage mit dieser exakten Kombination im Monat (Gewichtung)
        COUNT(*)               AS ANZ_TAGE

    FROM PROD_SDP_AL_PERS.AL_DIV_SHARED.DIV_TBL_HIERARCHIE

    WHERE VTWS IN ('11017', '11028', '11025', '11020', '11021', '11018', '11019')
      AND JAHR > '2016'

    GROUP BY
        -- Zeitdimension
        JAHR, MONAT,
        -- Vermittler
        BO, BO_NAME, BO_VVH, BO_SEIT, PA_ANTEIL,
        -- ZKTO1-8
        ZKTO1, ZKTO1_NAME, ZKTO1_VVH,
        ZKTO2, ZKTO2_NAME, ZKTO2_VVH,
        ZKTO3, ZKTO3_NAME, ZKTO3_VVH,
        ZKTO4, ZKTO4_NAME, ZKTO4_VVH,
        ZKTO5, ZKTO5_NAME, ZKTO5_VVH,
        ZKTO6, ZKTO6_NAME, ZKTO6_VVH,
        ZKTO7, ZKTO7_NAME, ZKTO7_VVH,
        ZKTO8, ZKTO8_NAME, ZKTO8_VVH,
        -- Eingangskontakt
        EKTO, EKTO_VVH, EKTO_NAME, EKTO_SEIT, EKTO_ALTVERS,
        -- Sonstige
        INSP, QUER,
        MKTO, MKTO_NAME, MKTO_VVH,
        MSB, MSB_NAME,
        MAKLCAT, MAKSEG, MAKSEG_TXT,
        VERBAND, MITGLNUMMER,
        -- IBN / IBSN
        IBN, IBN_NAME, IBSN, IBSN_NAME,
        FUEHR2_IBN, FUEHR2_IBN_NAME,
        FUEHR1_IBN, FUEHR1_IBN_NAME,
        FUEHR2_IBSN, FUEHR2_IBSN_NAME,
        FUEHR1_IBSN, FUEHR1_IBSN_NAME,
        VTWEG_IBN, VTWEG_IBSN,
        -- MBV
        MBV, MBV_NAME, MBV_VVH,
        FUEHR2_MBV, FUEHR2_MBV_NAME,
        FUEHR1_MBV, FUEHR1_MBV_NAME,
        VTWEG_MBV,
        -- MBKV
        MBKV, MBKV_NAME,
        FUEHR2_MBKV, FUEHR2_MBKV_NAME,
        FUEHR1_MBKV, FUEHR1_MBKV_NAME,
        VTWEG_MBKV,
        -- OEDSP
        OEDSP, OEDSP_NAME,
        FUEHR2_OEDSP, FUEHR2_OEDSP_NAME,
        FUEHR1_OEDSP, FUEHR1_OEDSP_NAME,
        VTWEG_OEDSP,
        -- MBCEB
        MBCEB, MBCEB_NAME,
        FUEHR2_MBCEB, FUEHR2_MBCEB_NAME,
        FUEHR1_MBCEB, FUEHR1_MBCEB_NAME,
        VTWEG_MBCEB,
        -- MBKI
        MBKI, MBKI_NAME,
        FUEHR2_MBKI, FUEHR2_MBKI_NAME,
        FUEHR1_MBKI, FUEHR1_MBKI_NAME,
        VTWEG_MBKI,
        -- KAMKV
        KAMKV, KAMKV_NAME,
        FUEHR2_KAMKV, FUEHR2_KAMKV_NAME,
        FUEHR1_KAMKV, FUEHR1_KAMKV_NAME,
        VTWEG_KAMKV,
        -- KAMV
        KAMV, KAMV_NAME,
        FUEHR2_KAMV, FUEHR2_KAMV_NAME,
        FUEHR1_KAMV, FUEHR1_KAMV_NAME,
        VTWEG_KAMV,
        -- MBK
        MBK, MBK_NAME, MBK_VVH,
        FUEHR2_MBK, FUEHR2_MBK_NAME,
        FUEHR1_MBK, FUEHR1_MBK_NAME,
        VTWEG_MBK,
        -- MRK
        MRK, MRK_NAME,
        FUEHR2_MRK, FUEHR2_MRK_NAME,
        FUEHR1_MRK, FUEHR1_MRK_NAME,
        VTWEG_MRK,
        -- MBPAV
        MBPAV, MBPAV_NAME,
        FUEHR2_MBPAV, FUEHR2_MBPAV_NAME,
        FUEHR1_MBPAV, FUEHR1_MBPAV_NAME,
        VTWEG_MBPAV,
        -- Führungskräfte / Betriebsleiter
        BETRLI, BETRLI_NAME, BETRLI_VVH,
        FUEHR2, FUEHR2_NAME, FUEHR2_TEXT,
        FUEHR1, FUEHR1_NAME, FUEHR1_TEXT,
        -- Vertriebsweg
        VTWS, VKL, VKL_NAME, VKL_VVH,
        -- Weitere Rollen
        AM, AM_NAME,
        FBP, FBP_NAME,
        FBCOR, FBCOR_NAME,
        MAB, MAB_NAME,
        RGB, RGB_NAME,
        FBOED, FBOED_NAME,
        FBCEB, FBCEB_NAME,
        LMAB, LMAB_NAME
)

SELECT * EXCLUDE (ANZ_TAGE)
FROM gezaehlt
-- Pro BO/JAHR/MONAT: nur die Kombination mit den meisten Tagen behalten.
-- Bei Gleichstand gewinnt die Gruppe, die zuerst gefunden wird (Snowflake-intern deterministisch).
QUALIFY ROW_NUMBER() OVER (
    PARTITION BY BO, JAHR, MONAT
    ORDER BY ANZ_TAGE DESC
) = 1;
