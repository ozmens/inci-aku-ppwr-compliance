# Phase F Migration Discrepancies

Count: 5

- DISC-0001 [WARN] CONFIGURATION_COUNT `LEVEL2_VS_GOLDEN` configuration_count: golden=247 secondary=293 → KEEP_GOLDEN (Level-1 Golden Register is authoritative for final configuration identity)
- DISC-0002 [INFO] SOURCE_FILE `Mamul Ambalaj Bileşen Listesi-30.07.2026(6).xlsx` role: golden='LEVEL_1_AUTHORITATIVE' secondary='LEVEL_3_STARTER_OPERATIONAL' → LINEAGE_ONLY (Operational source retained for lineage; not used to redefine final configs)
- DISC-0003 [INFO] SOURCE_FILE `Endüstriyel Ambalaj Miktarı Çalışması(5).xlsx` role: golden='LEVEL_1_AUTHORITATIVE' secondary='LEVEL_3_INDUSTRIAL_OPERATIONAL' → LINEAGE_ONLY (Operational source retained for lineage; not used to redefine final configs)
- DISC-0004 [INFO] SOURCE_FILE `Yüklemede Kullanılan Malzeme Ağırlıkları-2(7).xlsx` role: golden='LEVEL_1_AUTHORITATIVE' secondary='LEVEL_3_CONTAINER_OPERATIONAL' → LINEAGE_ONLY (Operational source retained for lineage; not used to redefine final configs)
- DISC-0005 [INFO] EVIDENCE_ARCHIVE `TEDARİKÇİLER.7z` sha256: golden='' secondary='b65c13a121d8236d64e241cf0f35343734557cd1369855ccdf68d86f3f4810db' → INVENTORIED_ONLY (Archive not modified. Component-to-document links not guessed. Drawings/photos remain PENDING.)
