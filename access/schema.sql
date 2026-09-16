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
-- cannot start with a leading space. "Blu-ray.com URL" and "Blu-ray.com
-- Title" -- both taken verbatim from the Excel/Calc canonical header list
-- in the handoff section 7.2 -- violate this (the period), so they are
-- named "Blu-ray com URL" and "Blu-ray com Title" here instead (period
-- dropped, kept as a space, matching every other multi-word field name in
-- this schema). Every other field name in this table is already
-- Access-safe and identical to its spreadsheet header text. See the
-- mapping table in access/README.md for the full Access-name <->
-- spreadsheet-header correspondence -- this matters for Phase 2, where
-- the resolver needs to write to the right Access field for each
-- spreadsheet column.

CREATE TABLE MediaCatalog (
    ID COUNTER PRIMARY KEY,
    [Inventory Number] TEXT(255),
    UPC TEXT(255),
    [Blu-ray com URL] TEXT(255),
    [Blu-ray com Title] TEXT(255),
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
