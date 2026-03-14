# Spreadsheet Guide — Tourismo Import Format for Heritage Listings

## Overview

Every site in an itinerary becomes a row in `upstate-listings-master.xlsx`.
Use `openpyxl` to append rows. Column order must be exact.

**File path:** `/mnt/user-data/outputs/upstate-listings-master.xlsx`

**If file doesn't exist:** Create with headers + rows.
**If file exists:** `load_workbook()`, find `ws.max_row`, append from `ws.max_row + 1`.

---

## Column Order (exact)

```
external_id | location_name | street_address | city | region_id |
latitude | longitude | subtitle | short_description | description |
province | postal_code | country | phone_number | email | website |
website2 | instagram_link | facebook_link | twitter_link | youtube_link |
pinterest_link | tiktok_link | tripadvisor_link | google_mybusiness_link |
tags | show_on_web | show_on_app
```

---

## Field Rules

| Field | Rule |
|---|---|
| `external_id` | `upstate-[slugified-name]` — lowercase, hyphens, no special chars |
| `location_name` | Official name of the site |
| `region_id` | `[County Name] County` — e.g. `Saratoga County`, `Cayuga County` |
| `province` | Always `NY` |
| `country` | Always `US` |
| `show_on_web` | Always `true` |
| `show_on_app` | Always `true` |
| Social links | Include if findable; leave blank string if not |
| `website2` | DMO listing page or trail program page as secondary URL |

---

## Tag Format

Tags are a JSON string. Minimum three categories: Business Category + Business Type + one Experience.

```json
{
  "Business Category": "Experience",
  "Business Type": "National Park",
  "County": ["Saratoga County"],
  "Region": ["Capital-Saratoga Region"],
  "Trail": "Path Through History",
  "Experience": ["History & Heritage", "Guided Tour", "Family-Friendly"]
}
```

---

## Business Category Options (heritage sites)

- `Experience` — battlefields, historic sites, forts, outdoor heritage
- `Cultural Heritage` — museums, art sites, cultural centers, house museums
- `Nature` — parks with significant natural AND historical value

---

## Business Type Options (heritage sites)

| Type | When to use |
|---|---|
| `National Park` | NPS-managed units |
| `State Historic Site` | NYS OPRHP-managed sites |
| `National Historic Landmark` | NHL-designated properties (privately operated) |
| `Museum` | Standalone history or art museums |
| `Living History Museum` | Museums with costumed interpretation |
| `Historic House` | House museums, presidential sites |
| `Cultural Center` | Indigenous cultural centers |
| `Heritage Center` | UGRR, canal, ethnic heritage centers |
| `Historic District` | Walkable historic districts (e.g. Huguenot Street) |
| `Battlefield` | Revolutionary War, War of 1812, Civil War |
| `Fort` | Military forts, reconstructed or ruined |
| `Estate` | Gilded Age estates, presidential estates |
| `Great Camp` | Adirondack Great Camps |
| `Arts Colony` | Artist colonies and creative communities |

---

## Region Options

- `Capital-Saratoga Region`
- `Hudson Valley`
- `Catskills`
- `Mohawk Valley`
- `Central NY`
- `Finger Lakes`
- `Adirondacks`
- `North Country`
- `Southern Tier`
- `Western NY`

---

## Trail / Program Tags (use most specific applicable)

| Tag | Use for |
|---|---|
| `Path Through History` | Default for any NY heritage site not on a specific trail |
| `Women's Rights History Trail` | Sites connected to suffrage and women's rights |
| `Underground Railroad Heritage Trail` | UGRR-connected sites |
| `Erie Canalway Trail` | Erie Canal corridor sites |
| `Haudenosaunee Heritage Trail` | Indigenous Haudenosaunee sites |
| `Revolutionary War Sites` | Sites without a more specific trail affiliation |
| `Haunted History Trail of New York State` | Sites with haunted history connection |

Use an array for multiple trails: `"Trail": ["Path Through History", "Women's Rights History Trail"]`

---

## Experience Tags (for heritage sites)

- `History & Heritage` — primary tag; use on every heritage site
- `Guided Tour` — formal tour programs available
- `Self-Guided Tour` — self-guided options available
- `Living History` — costumed reenactment or living history programming
- `Hiking` — significant trail network on site
- `Scenic Views` — notable landscapes (Olana, Saratoga overlooks)
- `Educational` — strong interpretive/education programming
- `Family-Friendly` — particularly accessible for families
- `Indigenous Culture` — Haudenosaunee and other Indigenous heritage
- `Art & Culture` — museums, artists' sites, cultural institutions

---

## Content Writing Rules for Listing Rows

### `subtitle` (15 words or fewer)

Ground the site in a specific historical moment, person, or consequence.
Sound like something you'd say to a trusted friend.

### `short_description` (1–2 sentences)

Lead with the specific non-obvious thing. Not what the site is — what it reveals.

### `description` (80–150 words)

Structure: **What happened here -> What's there now -> Don't Miss**

---

## Deduplication Note

If a site already appears in the spreadsheet (same `external_id`), add a note in `website2`
rather than creating a duplicate row — or flag it for manual review. The `external_id` is
the unique key.
