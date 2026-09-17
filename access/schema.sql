-- MediaCatalog Access table schema (Jet/ACE SQL, Access 2019 / .accdb).
--
-- This is the canonical, hand-maintained source of truth for the table's
-- structure -- build_access_database.vbs executes this file verbatim via
-- CurrentDb.Execute to create the table, so keep both in sync.
--
-- One flat table, no parent/child hierarchy: box-set/child-record support
-- is explicitly deferred (see the project handoff, section 7.3).
--
-- Field types:
--   - Identifier-like fields (Inventory Number, UPC, the two URL columns,
--     IMDb ID) are TEXT, never a numeric type. This project has already
--     been bitten twice by numeric-string bugs in Excel and LibreOffice
--     Calc (leading zeros dropped from UPCs, a movie titled "1917" turned
--     into a number); Access has no equivalent silent auto-coercion on
--     assignment, so this isn't strictly required here for correctness,
--     but keeping identifiers as TEXT keeps the schema honest about what
--     these values are (opaque codes, not quantities) and matches every
--     other MediaCatalog front end.
--   - Year, Runtime, Blu-ray Year, Blu-ray Runtime, and Season are genuine
--     numeric quantities used for sorting/filtering, so they are proper
--     INTEGER fields here -- unlike Excel/Calc, Access doesn't silently
--     reinterpret a text field as a number, so there's no reason to store
--     these as text and lose real sort/filter/aggregate behavior.
--   - Physical Release Date is a real DATETIME field.
--   - Status / Error is MEMO (Long Text), since it can carry more than 255
--     characters once warnings are appended.
--   - Every other field is TEXT(255) (Short Text).
--
-- ID is a surrogate AutoNumber primary key; nothing in this phase enforces
-- uniqueness on Inventory Number or UPC (a box set may have several owned
-- copies, or no UPC at all -- see the handoff's box-set/collection notes).
--
-- Access field names cannot contain a period, "!", "`", "[", or "]", and
-- cannot start with a leading space. The canonical spreadsheet headers were
-- originally "Blu-ray.com URL" and "Blu-ray.com Title" (violating this on
-- the period), so this table used "Blu-ray com URL"/"Blu-ray com Title"
-- as Access-only field names with a separate display label carrying the
-- real header text. The spreadsheet headers were later renamed to
-- "Blu-ray URL"/"Blu-ray Title" (dropping ".com" entirely, matching
-- "Blu-ray Year"/"Blu-ray Runtime", which never had it) specifically so
-- the Access field names below could become identical to the spreadsheet
-- header text -- no separate Access-name mapping needed for these two
-- fields any more. Every field name in this table is Access-safe and
-- identical to its spreadsheet header text.

CREATE TABLE MediaCatalog (
    ID COUNTER PRIMARY KEY,
    [Inventory Number] TEXT(255),
    UPC TEXT(255),
    [Blu-ray URL] TEXT(255),
    [Blu-ray Title] TEXT(255),
    [IMDb URL] TEXT(255),
    [IMDb ID] TEXT(255),
    [IMDb Title] TEXT(255),
    Year INTEGER,
    Runtime INTEGER,
    [Title Type] TEXT(255),
    Season INTEGER,
    [Status / Error] MEMO,
    Studio TEXT(255),
    [Blu-ray Year] INTEGER,
    [Blu-ray Runtime] INTEGER,
    [Content Rating] TEXT(255),
    [Physical Release Date] DATETIME,
    [Disc Format] TEXT(255),
    [Video Codec] TEXT(255),
    Resolution TEXT(255),
    [Aspect Ratio] TEXT(255),
    [Disc Count / Capacities] TEXT(255)
);
