# Excel Power Query — Merge & Join Types

Power Query's **Merge** feature lets you combine two tables by matching on a common column — similar to SQL JOINs. This guide walks through each join type using real sample data.

---

## Sample Data

#### Table 1 — Policy Master

| Policy ID   | Insured Name   | Region | Portal Premium | Source Premium |
|-------------|----------------|--------|---------------|----------------|
| ZNA-937933  | David Harris   | NC     | 59,943.78     | 59,943.78      |
| ZNA-365005  | Karen Garcia   | GA     | 29,073.47     | 21,304.11      |
| ZNA-441737  | Sarah Anderson | GA     | 57,057.49     | 57,057.49      |
| ZNA-306899  | Linda Williams | CT     | 40,273.00     | 40,273.00      |
| **ZNA-999001**  | **Tom Baker**      | **TX** | **15,000.00** | **15,000.00**  |

#### Table 2 — Policy Status

| Policy ID   | Driver Name    | State | Effective Date | Portal Premium | Status   |
|-------------|----------------|-------|---------------|----------------|----------|
| ZNA-937933  | David Harris   | NC    | 45,423        | 59,943.78      | Inactive |
| ZNA-365005  | Karen Garcia   | GA    | 45,550        | 29,073.47      | Inactive |
| ZNA-441737  | Sarah Anderson | GA    | 45,580        | 57,057.49      | Pending  |
| ZNA-306899  | Linda Williams | CT    | 45,644        | 40,273.00      | Inactive |
| **ZNA-888002**  | **Jane Smith**     | **FL**| **45,700**    | **22,000.00**  | **Active** |

> **Key difference:** Table 1 has `ZNA-999001` (Tom Baker) with **no match** in Table 2. Table 2 has `ZNA-888002` (Jane Smith) with **no match** in Table 1. The remaining 4 rows match on **Policy ID**.

---

## Join Types Explained

#### Left Outer — All rows from Table 1, matching rows from Table 2

Keeps **every row from Table 1**. If a match exists in Table 2, its columns are filled in. If no match is found, Table 2 columns are filled with `null`.

**Use when:** Table 1 is your primary/master table and you want to enrich it with Table 2 data without losing any records.

| Policy ID   | Insured Name   | Region | Status   |
|-------------|----------------|--------|----------|
| ZNA-937933  | David Harris   | NC     | Inactive |
| ZNA-365005  | Karen Garcia   | GA     | Inactive |
| ZNA-441737  | Sarah Anderson | GA     | Pending  |
| ZNA-306899  | Linda Williams | CT     | Inactive |
| ZNA-999001  | Tom Baker      | TX     | **null** |

> Tom Baker is retained from Table 1 but has no Status — Jane Smith from Table 2 is excluded entirely.

---

#### Right Outer — All rows from Table 2, matching rows from Table 1

Keeps **every row from Table 2**. If a match exists in Table 1, its columns are filled in. If no match is found, Table 1 columns are filled with `null`.

**Use when:** Table 2 is the authoritative source and you want to retain all its records.

| Policy ID   | Insured Name   | Region | Status   |
|-------------|----------------|--------|----------|
| ZNA-937933  | David Harris   | NC     | Inactive |
| ZNA-365005  | Karen Garcia   | GA     | Inactive |
| ZNA-441737  | Sarah Anderson | GA     | Pending  |
| ZNA-306899  | Linda Williams | CT     | Inactive |
| ZNA-888002  | **null**       | **null** | Active |

> Jane Smith is retained from Table 2 but has no Insured Name or Region — Tom Baker from Table 1 is excluded entirely.

---

#### Full Outer — All rows from both tables

Keeps **every row from both tables**. Matching rows are combined; non-matching rows appear with `null` in the columns from the other table.

**Use when:** You need a complete picture and cannot afford to lose any records from either side.

| Policy ID   | Insured Name   | Region | Status   |
|-------------|----------------|--------|----------|
| ZNA-937933  | David Harris   | NC     | Inactive |
| ZNA-365005  | Karen Garcia   | GA     | Inactive |
| ZNA-441737  | Sarah Anderson | GA     | Pending  |
| ZNA-306899  | Linda Williams | CT     | Inactive |
| ZNA-999001  | Tom Baker      | TX     | **null** |
| ZNA-888002  | **null**       | **null** | Active |

> Both unmatched rows (Tom Baker and Jane Smith) are included. No data is lost.

---

#### Inner — Only rows that match in both tables

Keeps **only rows where the Policy ID exists in both** Table 1 and Table 2. Unmatched rows from either side are dropped.

**Use when:** You only care about records that are confirmed in both sources — the "clean overlap."

| Policy ID   | Insured Name   | Region | Status   |
|-------------|----------------|--------|----------|
| ZNA-937933  | David Harris   | NC     | Inactive |
| ZNA-365005  | Karen Garcia   | GA     | Inactive |
| ZNA-441737  | Sarah Anderson | GA     | Pending  |
| ZNA-306899  | Linda Williams | CT     | Inactive |

> Tom Baker (Table 1 only) and Jane Smith (Table 2 only) are both excluded.

---

#### Left Anti — Rows in Table 1 with NO match in Table 2

Keeps **only the rows from Table 1 that do NOT have a matching row in Table 2**. The opposite of Inner for the left side.

**Use when:** You want to find records that are missing from Table 2 — great for identifying gaps or orphaned records.

| Policy ID   | Insured Name | Region | Portal Premium |
|-------------|--------------|--------|---------------|
| ZNA-999001  | Tom Baker    | TX     | 15,000.00     |

> Only Tom Baker is returned — he exists in Table 1 but has no corresponding entry in Table 2.

---

#### Right Anti — Rows in Table 2 with NO match in Table 1

Keeps **only the rows from Table 2 that do NOT have a matching row in Table 1**. The mirror image of Left Anti.

**Use when:** You want to find records that exist in Table 2 but are absent from Table 1 — useful for spotting new or unregistered entries.

| Policy ID   | Driver Name | State | Status |
|-------------|-------------|-------|--------|
| ZNA-888002  | Jane Smith  | FL    | Active |

> Only Jane Smith is returned — she exists in Table 2 but has no corresponding entry in Table 1.

---

## Quick Reference

| Join Type   | Rows from Table 1 | Rows from Table 2 | Best For |
|-------------|:-----------------:|:-----------------:|----------|
| Left Outer  | All            | Matching only  | Enrich master table |
| Right Outer | Matching only  | All            | Preserve secondary table |
| Full Outer  | All            | All            | Complete union, no data loss |
| Inner       | Matching only  | Matching only  | Verified overlap only |
| Left Anti   | Non-matching   | None           | Find gaps in Table 2 |
| Right Anti  | None           | Non-matching   | Find gaps in Table 1 |

---

## How to Perform a Merge in Power Query

1. Open **Power Query Editor** (`Data → Get Data → Launch Power Query Editor`)
2. Select your base table (Table 1)
3. Go to **Home → Merge Queries**
4. Select the **Right table** from the dropdown (Table 2)
5. Click the matching column in each table preview to set the join key
6. Choose your **Join kind** from the six options
7. Click **OK**, then expand the merged column using the expand icon (⇔)

---

*For fuzzy matching (e.g., matching "David Harris" with "D. Harris"), check **Use fuzzy matching to perform the merge** before clicking OK.*

---

## ⚠️ Watch Out: Duplicate Keys

A common real-world gotcha — if your join key (e.g. **Policy ID**) appears **more than once** in either table, Power Query will multiply the rows, not consolidate them.

#### Example — Duplicate in Table 2

| Policy ID  | Status   |
|------------|----------|
| ZNA-937933 | Inactive |
| ZNA-937933 | Active   |

If `ZNA-937933` appears **twice** in Table 2 and you do a **Left Outer** join, David Harris will appear **twice** in your result — once per matching row:

| Policy ID  | Insured Name | Region | Status   |
|------------|--------------|--------|----------|
| ZNA-937933 | David Harris | NC     | Inactive |
| ZNA-937933 | David Harris | NC     | Active   |

This is not a bug — it is intentional SQL-style behaviour. But it will silently inflate your row count and skew any aggregations (sums, counts, averages) downstream.

#### How to prevent it

- **Before merging**, check for duplicates on your key column:
  - Select the key column → `Home → Keep Rows → Keep Duplicates` — if the result is empty, you are safe
- **Deduplicate first** if needed: `Home → Remove Rows → Remove Duplicates` on the key column
- **After merging**, compare the row count to your original Table 1 — an unexpected increase is a sign duplicates crept in
