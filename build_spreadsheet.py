#!/usr/bin/env python3
"""Build the upstate-listings-master.xlsx spreadsheet for the Turning Point itinerary."""

import json
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/home/user/250th/outputs/upstate-listings-master.xlsx"

HEADERS = [
    "external_id", "location_name", "street_address", "city", "region_id",
    "latitude", "longitude", "subtitle", "short_description", "description",
    "province", "postal_code", "country", "phone_number", "email", "website",
    "website2", "instagram_link", "facebook_link", "twitter_link", "youtube_link",
    "pinterest_link", "tiktok_link", "tripadvisor_link", "google_mybusiness_link",
    "tags", "show_on_web", "show_on_app"
]

HEADER_FILL = PatternFill("solid", fgColor="1A3557")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ROW_FILL_ODD = PatternFill("solid", fgColor="FFFFFF")
ROW_FILL_EVEN = PatternFill("solid", fgColor="EEF3FA")
BODY_FONT = Font(name="Arial", size=9)

COL_WIDTHS = {
    "external_id": 38, "location_name": 35, "street_address": 28,
    "city": 18, "region_id": 22, "latitude": 12, "longitude": 12,
    "subtitle": 48, "short_description": 55, "description": 80,
    "province": 8, "postal_code": 10, "country": 8,
    "phone_number": 16, "email": 28, "website": 45,
    "website2": 45, "instagram_link": 40, "facebook_link": 40,
    "twitter_link": 35, "youtube_link": 35, "pinterest_link": 35,
    "tiktok_link": 35, "tripadvisor_link": 40,
    "google_mybusiness_link": 40, "tags": 60,
    "show_on_web": 12, "show_on_app": 12
}


def slugify(name):
    return "upstate-" + re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')


def build_tags(category, btype, counties, regions, trails, experiences):
    return json.dumps({
        "Business Category": category,
        "Business Type": btype,
        "County": counties if isinstance(counties, list) else [counties],
        "Region": regions if isinstance(regions, list) else [regions],
        "Trail": trails,
        "Experience": experiences if isinstance(experiences, list) else [experiences]
    })


# Define all 10 sites from the itinerary
rows = [
    {
        "external_id": slugify("Schuyler Mansion State Historic Site"),
        "location_name": "Schuyler Mansion State Historic Site",
        "street_address": "32 Catherine Street",
        "city": "Albany",
        "region_id": "Albany County",
        "latitude": "42.6430",
        "longitude": "-73.7520",
        "subtitle": "Where Hamilton married into the revolution — and a Loyalist hatchet scarred the stairs",
        "short_description": "Philip Schuyler coordinated the northern frontier's defense from this Georgian mansion while Congress tried to replace him. Hamilton married Eliza Schuyler here in December 1780, and Loyalist raiders left hatchet marks in the front staircase that survive today.",
        "description": "Philip Schuyler built this Georgian mansion in 1765 on a bluff above the Hudson. From here he organized the scorched-earth retreat that slowed Burgoyne's advance to a crawl — and watched Congress give the resulting victory to Horatio Gates. The house hosted Washington, Franklin, and the December 1780 wedding of Alexander Hamilton and Elizabeth Schuyler. In 1781, Loyalist raiders broke through the front door in an attempt to kidnap Schuyler; the hatchet marks on the staircase railing survive. The rooms have been restored to their 18th-century appearance. Don't miss: the hatchet marks on the front staircase railing — scars from the 1781 Loyalist raid that nearly captured the general in his own home.",
        "province": "NY",
        "postal_code": "12202",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://parks.ny.gov/historic-sites/schuyler-mansion",
        "website2": "https://www.iloveny.com/listing/schuyler-mansion-state-historic-site/3795/",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Cultural Heritage", "State Historic Site", "Albany County",
                           "Capital-Saratoga Region", "Revolutionary War Sites",
                           ["History & Heritage", "Guided Tour", "Educational"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Saratoga National Historical Park"),
        "location_name": "Saratoga National Historical Park",
        "street_address": "648 Route 32",
        "city": "Stillwater",
        "region_id": "Saratoga County",
        "latitude": "42.9981",
        "longitude": "-73.6411",
        "subtitle": "Where 5,895 British troops surrendered — and France decided to back a revolution",
        "short_description": "The battles of Freeman's Farm and Bemis Heights destroyed Burgoyne's army in September and October 1777. Benedict Arnold's unauthorized charge broke the British line — and a musket ball broke his leg. The Boot Monument honors the leg but leaves Arnold's name off the pedestal.",
        "description": "Between September 19 and October 7, 1777, two battles on these fields destroyed Burgoyne's army and changed the course of the war. At Freeman's Farm, Daniel Morgan's riflemen broke the British center. Three weeks later, Benedict Arnold charged into the second battle, rallied the American left, and took a musket ball in his leg. His charge broke the British line. Burgoyne surrendered 5,895 men on October 17 — the first capitulation of an entire British army. France recognized the United States within months. The 9-mile battlefield tour road passes through the actual terrain. Don't miss: the Boot Monument at Stop 7 — a carved bas-relief of a boot honoring 'the most brilliant soldier of the Continental Army,' with Arnold's name deliberately left off the pedestal.",
        "province": "NY",
        "postal_code": "12170",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.nps.gov/sara",
        "website2": "https://america250.org",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "National Park", "Saratoga County",
                           "Capital-Saratoga Region", "Revolutionary War Sites",
                           ["History & Heritage", "Self-Guided Tour", "Scenic Views", "Educational", "Family-Friendly"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Schuyler House Saratoga"),
        "location_name": "Schuyler House (Saratoga NHP)",
        "street_address": "Route 4",
        "city": "Schuylerville",
        "region_id": "Saratoga County",
        "latitude": "43.1003",
        "longitude": "-73.5831",
        "subtitle": "Burgoyne burned it on his retreat — Schuyler rebuilt it within a year",
        "short_description": "After Burgoyne retreated north from the battlefield, he burned Schuyler's country estate. Schuyler rebuilt this house within a year — a declaration that the Schuylers weren't leaving. The site sits near the surrender field where Burgoyne's army laid down their arms.",
        "description": "When Burgoyne's retreating army burned Philip Schuyler's country estate in October 1777, it was an act of spite by a losing general. Schuyler rebuilt this house within a year on the same foundation — equal parts stubbornness and strategy. The house sits a few miles from the field where Burgoyne surrendered 5,895 men, and its reconstruction was as much a political statement as a personal one: the Schuylers owned this valley, and they intended to stay. The interior reflects Schuyler's wealth and the domestic life of one of the revolution's architects. Don't miss: the view from the back of the house toward the Hudson — the same view Schuyler saw when he chose to rebuild on scorched ground.",
        "province": "NY",
        "postal_code": "12871",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.nps.gov/sara",
        "website2": "",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Cultural Heritage", "Historic House", "Saratoga County",
                           "Capital-Saratoga Region", "Revolutionary War Sites",
                           ["History & Heritage", "Guided Tour"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Saratoga Monument"),
        "location_name": "Saratoga Monument",
        "street_address": "1 Burgoyne Street",
        "city": "Schuylerville",
        "region_id": "Saratoga County",
        "latitude": "43.0997",
        "longitude": "-73.5814",
        "subtitle": "Four niches for four generals — one has been empty since 1883",
        "short_description": "The 155-foot obelisk has four niches for the American generals at Saratoga. Three hold statues: Schuyler, Gates, Morgan. The fourth — Arnold's — has been empty since the monument was dedicated in 1883. You can climb the interior staircase for a view across the entire campaign landscape.",
        "description": "Dedicated in 1883, this 155-foot obelisk in the center of Schuylerville commemorates the American victory at Saratoga with four niches — one for each commanding general. Three hold bronze statues of Schuyler, Gates, and Morgan. The fourth niche, meant for Benedict Arnold, stands deliberately empty. Arnold's charge at Bemis Heights broke the British line and arguably won the battle, but his subsequent treason made him unmemorable in bronze. The empty niche is the most eloquent commentary on treason and heroism in American public art. Climb the interior staircase to the top for a panoramic view of the Hudson Valley that maps the entire 1777 campaign. Don't miss: the empty fourth niche — stand in front of it and consider what it means that a nation would carve a space for a hero it couldn't name.",
        "province": "NY",
        "postal_code": "12871",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.nps.gov/sara",
        "website2": "",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "Battlefield", "Saratoga County",
                           "Capital-Saratoga Region", "Revolutionary War Sites",
                           ["History & Heritage", "Scenic Views"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Fort Stanwix National Monument"),
        "location_name": "Fort Stanwix National Monument",
        "street_address": "112 East Park Street",
        "city": "Rome",
        "region_id": "Oneida County",
        "latitude": "43.2106",
        "longitude": "-75.4557",
        "subtitle": "550 men held this fort for 21 days — and stopped a British army cold",
        "short_description": "Colonel Peter Gansevoort held this frontier fort against a British force three times his size for 21 days in August 1777. The western arm of Burgoyne's strategy to split the colonies died here, at a bend in the Mohawk River that most travelers drive past.",
        "description": "In August 1777, Colonel Peter Gansevoort held this frontier fort with 550 men against Barry St. Leger's force of 1,700 British regulars, Loyalists, and Haudenosaunee allies for 21 days — long enough for Benedict Arnold's relief column to force a retreat. The western arm of Burgoyne's plan to split the colonies in two died here, at a bend in the Mohawk River. The NPS has reconstructed the fort to its 1758 appearance; uniformed rangers know the siege hour by hour. But the fort's significance stretches beyond 1777: the 1768 Treaty of Fort Stanwix, signed here, ceded vast Haudenosaunee territories and set the stage for every conflict that followed. Don't miss: the 1768 Treaty exhibit — the document that redrew the boundary between colonial settlement and Haudenosaunee territory.",
        "province": "NY",
        "postal_code": "13440",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.nps.gov/fost",
        "website2": "",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "National Park", "Oneida County",
                           "Mohawk Valley",
                           ["Revolutionary War Sites", "Haudenosaunee Heritage Trail"],
                           ["History & Heritage", "Living History", "Guided Tour", "Educational", "Family-Friendly"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Oriskany Battlefield State Historic Site"),
        "location_name": "Oriskany Battlefield State Historic Site",
        "street_address": "7801 State Route 69",
        "city": "Oriskany",
        "region_id": "Oneida County",
        "latitude": "43.1570",
        "longitude": "-75.3730",
        "subtitle": "The bloodiest battle of the Revolution by casualty rate — and the day the Confederacy fractured",
        "short_description": "On August 6, 1777, General Herkimer's relief column walked into an ambush in this ravine. The fighting was hand-to-hand. Oneida fought Mohawk. When it was over, the Haudenosaunee Confederacy's centuries-old unity was broken along lines that never fully healed.",
        "description": "On August 6, 1777, a relief column of 800 Oneida warriors and Tryon County militia marching to break the siege of Fort Stanwix walked into an ambush in a ravine beside Oriskany Creek. The fighting was hand-to-hand. Neighbors killed neighbors. Oneida warriors fought Mohawk warriors in what was effectively a Haudenosaunee civil war. General Nicholas Herkimer took a musket ball through the knee, had his men prop him against a beech tree, lit his pipe, and directed the battle for six hours while bleeding to death. The casualty rate was the highest of any Revolutionary War battle. The monument on the hill marks where Herkimer sat. Don't miss: stand at the edge of the ravine where the ambush began — the terrain explains the battle better than any exhibit.",
        "province": "NY",
        "postal_code": "13424",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://parks.ny.gov/historic-sites/oriskany-battlefield",
        "website2": "",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "Battlefield", "Oneida County",
                           "Mohawk Valley", "Revolutionary War Sites",
                           ["History & Heritage", "Self-Guided Tour", "Indigenous Culture"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Herkimer Home State Historic Site"),
        "location_name": "Herkimer Home State Historic Site",
        "street_address": "200 State Route 169",
        "city": "Little Falls",
        "region_id": "Herkimer County",
        "latitude": "43.0120",
        "longitude": "-74.8840",
        "subtitle": "The general who commanded from a beech tree — and died reading his Bible in German",
        "short_description": "Nicholas Herkimer was carried from Oriskany to this house on the Mohawk River, where a botched amputation killed him ten days later. He reportedly read his Bible aloud in German — his first language — as he died. The house is a fine Georgian brick home in a county of frame farmhouses.",
        "description": "Nicholas Herkimer was the wealthiest man in the Mohawk Valley and chose to march anyway. After taking a musket ball through the knee at Oriskany, he was carried to this Georgian brick house on the Mohawk River. A Continental Army surgeon amputated his shattered leg ten days after the battle. The amputation went wrong. Herkimer bled to death on August 16, 1777, reportedly reading his Bible aloud in German — his first language — to his family. The house stands in a landscape of frame farmhouses, a reminder of Herkimer's status and his sacrifice. Little Falls (population 4,500) is a Mohawk Valley mill town that rewards a slow drive through. Don't miss: the room where Herkimer died — the family Bible and bed placement are based on period accounts.",
        "province": "NY",
        "postal_code": "13365",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://parks.ny.gov/historic-sites/herkimer-home",
        "website2": "https://mohawkvalleyhistory.com",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Cultural Heritage", "State Historic Site", "Herkimer County",
                           "Mohawk Valley", "Revolutionary War Sites",
                           ["History & Heritage", "Guided Tour"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Old Stone Fort Museum"),
        "location_name": "Old Stone Fort Museum",
        "street_address": "145 Fort Road",
        "city": "Schoharie",
        "region_id": "Schoharie County",
        "latitude": "42.6660",
        "longitude": "-74.3120",
        "subtitle": "A 1772 church fortified for war — holding the Palatine German story no one else tells",
        "short_description": "Schoharie is a county seat of 1,200 people. This 1772 Reformed church, fortified during the Revolution and raided by Joseph Brant in 1780, holds the story of the Palatine German migration of 1710 — 3,000 Rhineland refugees who became the backbone of the Mohawk Valley militia.",
        "description": "Schoharie is a county seat of 1,200 people, and the Old Stone Fort Museum on the edge of town tells a story you'll find nowhere else. In 1710, three thousand Palatine German refugees from the Rhineland settled the Schoharie and Mohawk Valleys. Their descendants became the backbone of the frontier militia that fought at Oriskany and defended the valley throughout the Revolution. This 1772 Reformed church was fortified during the war and survived a 1780 raid by Joseph Brant and Sir John Johnson. The collection inside is local, personal, and irreplaceable: ledgers, tools, family Bibles in German script, and a genealogy archive that traces Palatine families back to 1710. Don't miss: the Palatine German genealogy collection — if your family came through the Schoharie Valley, the records here go back three centuries.",
        "province": "NY",
        "postal_code": "12157",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.theoldstonefort.org",
        "website2": "",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Cultural Heritage", "Museum", "Schoharie County",
                           "Capital-Saratoga Region",
                           ["Revolutionary War Sites", "Path Through History"],
                           ["History & Heritage", "Self-Guided Tour", "Educational"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Crown Point State Historic Site"),
        "location_name": "Crown Point State Historic Site",
        "street_address": "739 Bridge Road",
        "city": "Crown Point",
        "region_id": "Essex County",
        "latitude": "43.9430",
        "longitude": "-73.4250",
        "subtitle": "Ruins of two great forts on a limestone promontory — and almost nobody comes here",
        "short_description": "Two days after Allen took Ticonderoga, Seth Warner captured Crown Point — and 111 more cannons. The ruins of the massive 1759 British fort and the earlier French Fort St. Frederic form one of the most impressive military landscapes in North America. Crown Point (population 2,000) is an RTN destination.",
        "description": "Two days after Ethan Allen captured Fort Ticonderoga, Seth Warner took Crown Point and its 111 additional cannons. The ruins here are older and more imposing than Ticonderoga's: the massive stone walls of the 1759 British fort — built to hold 4,000 troops — and the remains of the earlier French Fort St. Frederic form one of the most impressive colonial military landscapes in North America. Crown Point (population 2,000) sits on a limestone promontory jutting into Lake Champlain, and on a clear day the view reaches into Vermont and Canada. Almost nobody comes here, which is exactly why you should. The scale of the crumbling walls against the lake is the most dramatic physical remnant of the colonial wars in New York. Don't miss: the ruins of Fort St. Frederic — the French fortification whose walls crumble directly into Lake Champlain.",
        "province": "NY",
        "postal_code": "12928",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://parks.ny.gov/historic-sites/crown-point",
        "website2": "https://visitadirondacks.com",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "Fort", "Essex County",
                           "Adirondacks", "Revolutionary War Sites",
                           ["History & Heritage", "Self-Guided Tour", "Scenic Views", "Hiking"]),
        "show_on_web": "true",
        "show_on_app": "true"
    },
    {
        "external_id": slugify("Fort Ticonderoga"),
        "location_name": "Fort Ticonderoga",
        "street_address": "102 Fort Ti Road",
        "city": "Ticonderoga",
        "region_id": "Essex County",
        "latitude": "43.8391",
        "longitude": "-73.3894",
        "subtitle": "America's first offensive victory — 59 cannons that ended the siege of Boston",
        "short_description": "Ethan Allen arrived with 83 men and no artillery plan. He left with the cannons that Henry Knox dragged 300 miles on ox sleds to end the siege of Boston. Fort Ticonderoga's REAL TIME REVOLUTION program — immersive and running through 2027 — is built for the 250th anniversary.",
        "description": "On the night of May 10, 1775, Ethan Allen and 83 Green Mountain Boys crossed Lake Champlain in the dark and captured the British garrison in what one soldier called 'a confused melee of shouting.' It was America's first offensive military victory, and it mattered less for the symbolism than for the 59 cannons Henry Knox dragged 300 miles on ox sleds to Dorchester Heights the following winter, ending the siege of Boston. The fort has been restored to its 18th-century appearance by the Mars family since 1908 and is the most thoroughly interpreted colonial military site in the country. The King's Garden, between the fort and the lake, is one of the most beautiful cultivated spaces in the Adirondacks. Don't miss: the view from the ramparts north across Lake Champlain to Mount Defiance — the high ground the British seized in 1777 to force the American evacuation.",
        "province": "NY",
        "postal_code": "12883",
        "country": "US",
        "phone_number": "",
        "email": "",
        "website": "https://www.fortticonderoga.org",
        "website2": "https://america250.org",
        "instagram_link": "",
        "facebook_link": "",
        "twitter_link": "",
        "youtube_link": "",
        "pinterest_link": "",
        "tiktok_link": "",
        "tripadvisor_link": "",
        "google_mybusiness_link": "",
        "tags": build_tags("Experience", "Fort", "Essex County",
                           "Adirondacks", "Revolutionary War Sites",
                           ["History & Heritage", "Living History", "Guided Tour", "Scenic Views", "Educational", "Family-Friendly"]),
        "show_on_web": "true",
        "show_on_app": "true"
    }
]

# Build the workbook
wb = Workbook()
ws = wb.active
ws.title = "Listings"

# Write headers
for col_idx, header in enumerate(HEADERS, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

ws.row_dimensions[1].height = 20

# Set column widths
for col_idx, header in enumerate(HEADERS, 1):
    ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(header, 20)

ws.freeze_panes = "A2"

# Write rows
for row_offset, row_data in enumerate(rows):
    row_num = 2 + row_offset
    fill = ROW_FILL_ODD if row_num % 2 == 1 else ROW_FILL_EVEN
    for col_idx, header in enumerate(HEADERS, 1):
        val = row_data.get(header, "")
        cell = ws.cell(row=row_num, column=col_idx, value=val)
        cell.font = BODY_FONT
        cell.fill = fill
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[row_num].height = 60

wb.save(OUTPUT_PATH)
print(f"Created {OUTPUT_PATH} with {len(rows)} rows")
