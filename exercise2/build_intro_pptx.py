"""Build the Exercise 2 intro deck: Intro_to_Coordinates.pptx

Run inside the mapping_kids conda env:
    python make_intro_maps.py      (only needed once, makes images/)
    python build_intro_pptx.py

The drawing helpers and colours are shared with the Exercise 1 deck, so both
decks look the same. Every diagram is drawn with plain PowerPoint shapes.
"""
import sys
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx import Presentation

HERE = Path(__file__).parent
# Put exercise1 first on the path, so the import below finds its build_intro_pptx.py
sys.path.insert(0, str(HERE.parent / "exercise1"))
from build_intro_pptx import (  # noqa: E402  (the Exercise 1 deck's helpers)
    AMBER, AMBER_TINT, BLUE, BLUE_TINT, FONT, GREEN, GREEN_TINT, GREY_TINT, MUTED, NAVY,
    RED, TEXT, WHITE, add_arrow, add_badge, add_caption, add_card, add_dot, add_notes,
    add_picture, add_polyline, add_shape, add_text, add_title, fun_fact, new_slide,
    number_dot, word_of_the_day,
)

IMAGES = HERE / "images"
OUTPUT = HERE / "Intro_to_Coordinates.pptx"
LIGHT_GREY = RGBColor(0xD1, 0xD5, 0xDB)

# One colour per type of place, used on the "sorting" and "legend" slides
TYPE_COLOURS = [("Monument", RED), ("Palace", AMBER), ("Temple", BLUE), ("Nature", GREEN)]


# ---------- helpers used only in this deck ----------

def add_globe(slide, cx, cy, r):
    """A blue circle with a thin outline, the starting point for both globe drawings."""
    globe = add_shape(slide, MSO_SHAPE.OVAL, cx - r, cy - r, 2 * r, 2 * r, BLUE_TINT, line=BLUE)
    globe.line.width = Pt(2)
    return globe


def add_ring(slide, x, y, w, h, colour, width=2):
    """An outline-only oval, used for the curved longitude lines."""
    ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    ring.fill.background()
    ring.line.color.rgb = colour
    ring.line.width = Pt(width)
    ring.shadow.inherit = False
    return ring


def add_table(slide, x, y, w, rows, col_widths):
    """A simple table. The first row is the header."""
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y),
                                   Inches(w), Inches(0.5 * len(rows)))
    table = shape.table
    for c, width in enumerate(col_widths):
        table.columns[c].width = Inches(width)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if r == 0 else (WHITE if r % 2 else GREY_TINT)
            cell.text = value
            run = cell.text_frame.paragraphs[0].runs[0]
            run.font.name = FONT
            run.font.size = Pt(16)
            run.font.bold = r == 0
            run.font.color.rgb = WHITE if r == 0 else TEXT
    return shape


# ---------- the slides ----------

def slide_title(prs):
    s = new_slide(prs, dark=True)
    add_text(s, 0.8, 1.5, 6.2, 0.5, "MAPPING WITH KIDS  |  BEFORE EXERCISE 2",
             size=16, bold=True, color=AMBER)
    add_text(s, 0.8, 2.1, 6.2, 2.0, "Every Place Has an Address", size=48, bold=True,
             color=WHITE)
    add_text(s, 0.8, 4.2, 6.0, 1.2, "Latitude, longitude, and turning a table into a map",
             size=24, color=RGBColor(0xCB, 0xD5, 0xE1))
    add_picture(s, HERE / "screenshot" / "s2.png", 7.3, 2.0, w=5.4)
    add_notes(s, "Welcome back, mapmakers. Last time we drew the outline of India. Today we "
                 "learn the secret address that every spot on Earth has, and we use it to "
                 "put famous places of India on our map.")


def slide_journey(prs):
    s = new_slide(prs)
    add_title(s, "Today's journey")
    steps = [("1", "Addresses on Earth", BLUE), ("2", "Latitude", BLUE),
             ("3", "Longitude", BLUE), ("4", "From table to dots", GREEN),
             ("5", "Colours and labels", AMBER), ("6", "Distance and buffers", RED)]
    x = 0.75
    for i, (num, label, colour) in enumerate(steps):
        number = add_shape(s, MSO_SHAPE.OVAL, x + 0.45, 2.6, 1.1, 1.1, colour)
        tf = number.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = num
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.name = FONT
        run.font.color.rgb = WHITE
        add_text(s, x, 3.9, 2.0, 1.0, label, size=18, bold=True, align=PP_ALIGN.CENTER)
        if i < len(steps) - 1:
            add_arrow(s, x + 1.65, 3.15, x + 2.25, 3.15, MUTED, width=2)
        x += 2.0
    add_text(s, 0.75, 5.5, 11.8, 0.8,
             "At the end you will put 16 famous places of India on your own map.",
             size=22, color=MUTED, align=PP_ALIGN.CENTER)
    add_notes(s, "Steps 1 to 3 are about the numbers that tell us where a place is. Steps 4 "
                 "to 6 are the GIS skills you will use in Exercise 2. Ask: who has visited "
                 "a famous place in India? Which one?")


def slide_two_addresses(prs):
    s = new_slide(prs)
    add_title(s, "Every place has two addresses")
    add_card(s, 0.6, 1.6, 5.6, 2.6, "Postal address",
             ["Vidhana Soudha", "Ambedkar Veedhi", "Bengaluru 560001"], GREY_TINT, MUTED,
             size=20)
    add_card(s, 7.0, 1.6, 5.7, 2.6, "Earth address",
             ["Latitude: 12.98 N", "Longitude: 77.59 E"], BLUE_TINT, BLUE, size=24)
    add_arrow(s, 6.3, 2.9, 6.9, 2.9, MUTED, width=3)
    add_text(s, 0.6, 4.5, 6.4, 2.0,
             "A postal address only works if you know the street names. "
             "The Earth address works anywhere: in a city, a forest, or the middle of "
             "the sea.", size=20)
    word_of_the_day(s, 7.0, 4.5, 5.7, 1.9, "Coordinates",
                    "the two numbers, latitude and longitude, that pin down any spot on Earth.")
    add_notes(s, "Every house, school and shop has a postal address. But a postman cannot "
                 "deliver a letter to the middle of a forest. Scientists, pilots and ship "
                 "captains use coordinates instead. Bengaluru is about 13 degrees north and "
                 "77.6 degrees east. Your phone's map app uses these numbers all the time.")


def slide_latitude(prs):
    s = new_slide(prs)
    add_title(s, "Latitude: how far north or south")
    cx, cy, r = 3.4, 4.25, 2.55
    add_globe(s, cx, cy, r)
    # sin and cos of 30 and 60 degrees, so the lines sit at the right height
    lines = [(0.0, 1.0, "Equator  0", RED), (0.5, 0.866, "30 N", BLUE),
             (0.866, 0.5, "60 N", BLUE)]
    for sin_lat, cos_lat, label, colour in lines:
        y = cy - r * sin_lat
        half = r * cos_lat
        add_polyline(s, [(cx - half, y), (cx + half, y)], colour,
                     width=4 if colour == RED else 2)
        add_text(s, cx + half + 0.1, y - 0.25, 2.0, 0.5, label, size=16, bold=True,
                 color=colour)
    add_dot(s, cx, cy - r, 0.3, NAVY)
    add_text(s, cx - 2.65, cy - r - 0.35, 2.4, 0.5, "North Pole  90 N", size=16, bold=True,
             color=NAVY, align=PP_ALIGN.RIGHT)
    add_text(s, 7.2, 1.6, 5.5, 2.2,
             ["Lines that go across, like the rungs of a ladder",
              "Start at 0 at the equator, the middle of the Earth",
              "Go up to 90 at the North Pole and the South Pole"],
             size=20, bullets=True)
    fun_fact(s, 7.2, 4.3, 5.5, 1.9,
             "All of India is north of the equator. Its southern tip, Kanyakumari, is at "
             "about 8 N.")
    add_notes(s, "Latitude lines run east-west around the Earth, but they measure how far "
                 "north or south you are. A trick to remember: LAT-itude lines lie FLAT. "
                 "The equator is 0. Places north of it are N, places south of it are S. "
                 "Delhi is at about 29 N, Bengaluru at about 13 N, so Delhi is further north.")


def slide_longitude(prs):
    s = new_slide(prs)
    add_title(s, "Longitude: how far east or west")
    cx, cy, r = 3.4, 4.25, 2.55
    add_globe(s, cx, cy, r)
    # Meridians seen from the side look like narrower and narrower ovals
    for half_width in (0.5, 0.866):
        w = 2 * r * half_width
        add_ring(s, cx - w / 2, cy - r, w, 2 * r, BLUE)
    add_polyline(s, [(cx, cy - r), (cx, cy + r)], RED, width=4)
    add_text(s, cx - 1.6, cy + r + 0.05, 3.2, 0.5, "Prime meridian  0", size=16, bold=True,
             color=RED, align=PP_ALIGN.CENTER)
    add_text(s, 7.2, 1.6, 5.5, 2.4,
             ["Lines that go from the North Pole to the South Pole",
              "Start at 0 at Greenwich, in London",
              "Go up to 180 degrees east and 180 degrees west"],
             size=20, bullets=True)
    add_card(s, 7.2, 4.3, 5.5, 1.9, "India on the globe",
             "India is east of Greenwich, between about 68 E and 97 E.", BLUE_TINT, BLUE,
             size=20)
    add_notes(s, "Longitude lines are called meridians. They all meet at the poles, like the "
                 "segments of an orange. People agreed long ago to call the line through "
                 "Greenwich in London zero. India is east of it, so Indian longitudes end "
                 "in E. Latitude plus longitude gives one exact spot, like a square on a "
                 "chess board.")


def slide_india_grid(prs):
    s = new_slide(prs)
    add_title(s, "India on the grid")
    add_picture(s, IMAGES / "india_grid.png", 0.6, 1.35, h=5.8)
    add_text(s, 7.0, 1.5, 5.8, 2.4,
             ["Each grey line is 5 degrees apart",
              "Read a place's latitude on the left",
              "Read its longitude along the bottom"],
             size=20, bullets=True)
    add_card(s, 7.0, 3.9, 5.7, 2.4, "Tropic of Cancer",
             "This special line at 23.5 N runs right through the middle of India. "
             "Can you guess which states it crosses?", AMBER_TINT,
             RGBColor(0x92, 0x40, 0x0E), size=18)
    add_notes(s, "Point to Bengaluru: about 13 N, 77.6 E. The Tropic of Cancer crosses eight "
                 "states: Gujarat, Rajasthan, Madhya Pradesh, Chhattisgarh, Jharkhand, West "
                 "Bengal, Tripura and Mizoram. Ask: is Bengaluru north or south of the "
                 "Tropic of Cancer? (South.)")


def slide_find_the_place(prs):
    s = new_slide(prs)
    add_title(s, "Game: find the place")
    add_picture(s, IMAGES / "india_grid.png", 0.6, 1.35, h=5.8)
    add_text(s, 7.0, 1.5, 5.8, 0.6, "Where on the map are these places?", size=22,
             bold=True)
    mysteries = [("A", "27 N, 78 E", "A white marble tomb"),
                 ("B", "31.6 N, 74.9 E", "A temple with a free kitchen"),
                 ("C", "9.9 N, 78.1 E", "A temple with colourful towers")]
    y = 2.35
    for letter, coords, clue in mysteries:
        number_dot(s, 7.25, y + 0.35, letter, fill=AMBER)
        add_text(s, 7.75, y, 5.0, 0.5, coords, size=22, bold=True, color=NAVY)
        add_text(s, 7.75, y + 0.5, 5.0, 0.5, clue, size=18, color=MUTED)
        y += 1.35
    add_notes(s, "Students put a finger on the map where the lines cross. Answers: A is the "
                 "Taj Mahal in Agra. B is the Golden Temple in Amritsar. C is the Meenakshi "
                 "Temple in Madurai. All three are in today's data, so they will see them "
                 "again in Exercise 2.")


def slide_table(prs):
    s = new_slide(prs)
    add_title(s, "A table can hold places")
    rows = [["name", "type", "city", "latitude", "longitude"],
            ["Taj Mahal", "Monument", "Agra", "27.1751", "78.0421"],
            ["Hawa Mahal", "Palace", "Jaipur", "26.9239", "75.8267"],
            ["Golden Temple", "Temple", "Amritsar", "31.6200", "74.8765"],
            ["Kaziranga", "Nature", "Golaghat", "26.5775", "93.1711"]]
    add_table(s, 0.6, 1.6, 8.2, rows, [2.1, 1.6, 1.5, 1.5, 1.5])
    add_text(s, 0.6, 4.4, 8.2, 1.4,
             ["Each row is one place", "Each column is one fact about it",
              "Latitude and longitude tell the map where to put it"],
             size=20, bullets=True)
    word_of_the_day(s, 9.2, 1.6, 3.5, 2.6, "CSV",
                    "Comma-Separated Values. A plain text table with commas between the values.")
    add_notes(s, "This is part of the file you will use, famous_places.csv. It has 16 rows. "
                 "If you open it in Notepad you will see commas between every value: "
                 "Taj Mahal,Monument,Agra,27.1751,78.0421. Spreadsheets like Excel show the "
                 "same file as a neat grid.")


def slide_rows_to_dots(prs):
    s = new_slide(prs)
    add_title(s, "Each row becomes a dot")
    rows = [["name", "latitude", "longitude"],
            ["Delhi", "28.61", "77.21"],
            ["Kolkata", "22.57", "88.36"],
            ["Bengaluru", "12.97", "77.59"]]
    add_table(s, 0.6, 2.2, 5.0, rows, [2.0, 1.5, 1.5])
    add_arrow(s, 5.85, 3.2, 7.0, 3.2, GREEN, width=4)
    add_picture(s, IMAGES / "india_places.png", 7.3, 1.35, h=5.4)
    add_card(s, 0.6, 4.6, 5.0, 1.9, "Remember from Exercise 1",
             "Each place is a point: one dot, at one spot.", GREEN_TINT, GREEN, size=20)
    add_notes(s, "The GIS reads each row, finds the latitude and longitude, and puts a dot "
                 "there. Three rows give three dots. In Exercise 2 your table has 16 rows, "
                 "so you will get 16 dots. In Exercise 1 India was a polygon. Here each "
                 "city is so small on a map of India that a point is enough.")


def slide_sorting(prs):
    s = new_slide(prs)
    add_title(s, "Sorting places into groups")
    examples = {"Monument": ["Taj Mahal", "Qutub Minar"],
                "Palace": ["Hawa Mahal", "Mysore Palace"],
                "Temple": ["Golden Temple", "Konark Sun Temple"],
                "Nature": ["Kaziranga", "Valley of Flowers"]}
    x = 0.6
    for name, colour in TYPE_COLOURS:
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.6, 2.9, 3.0,
                  GREY_TINT).adjustments[0] = 0.06
        add_dot(s, x + 0.55, 2.2, 0.5, colour)
        add_text(s, x + 0.95, 1.95, 1.9, 0.5, name, size=22, bold=True, color=colour)
        add_text(s, x + 0.25, 2.8, 2.5, 1.6, examples[name], size=18, bullets=True)
        x += 3.1
    word_of_the_day(s, 0.6, 4.95, 7.4, 1.7, "Categorized style",
                    "giving each group of things on a map its own colour.")
    add_text(s, 8.4, 5.05, 4.3, 1.6,
             "One look at the colours tells you what kind of place each dot is.",
             size=18, color=MUTED)
    add_notes(s, "Our table has a type column with four groups. GIS software can colour the "
                 "dots by that column automatically. Your colours in GeoLibre may be "
                 "different from these, and that is fine. Ask: what group would your school "
                 "go in? What new group could we add?")


def slide_maps_know_facts(prs):
    s = new_slide(prs)
    add_title(s, "Maps that know facts")
    # a pretend map with one dot and a pop-up card
    add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 1.5, 6.2, 5.0,
              GREY_TINT).adjustments[0] = 0.04
    add_dot(s, 2.0, 5.2, 0.45, AMBER)
    add_badge(s, 2.35, 5.35, "Hawa Mahal", NAVY, size=14)
    add_arrow(s, 2.1, 4.95, 2.6, 4.3, MUTED, width=2)
    card = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 2.4, 1.8, 4.1, 2.45, WHITE, line=LIGHT_GREY)
    card.adjustments[0] = 0.06
    add_text(s, 2.6, 1.95, 3.8, 2.3,
             ["name: Hawa Mahal", "city: Jaipur", "type: Palace",
              "fun_fact: has 953 small windows"], size=16)
    add_text(s, 7.3, 1.6, 5.4, 0.5, "Two ways to see the facts", size=24, bold=True,
             color=NAVY)
    add_card(s, 7.3, 2.3, 5.4, 1.8, "Label",
             "Text on the map that comes from one column, such as the name.", AMBER_TINT,
             RGBColor(0x92, 0x40, 0x0E), size=18)
    add_card(s, 7.3, 4.35, 5.4, 1.8, "Attribute table",
             "The whole table behind the dots. Click a row and its dot lights up.", BLUE_TINT,
             BLUE, size=18)
    add_notes(s, "A dot on a paper map is just a dot. A dot in GIS remembers every fact from "
                 "its row in the table. A label writes one of those facts on the map. The "
                 "attribute table shows all of them. If you change a name in the table, the "
                 "label on the map changes too.")


def slide_crow_flies(prs):
    s = new_slide(prs)
    add_title(s, "As the crow flies")
    blr, mys = (5.6, 2.3), (1.3, 5.6)
    # the road wiggles through towns on the way
    road = [blr, (5.5, 3.3), (4.7, 3.6), (4.3, 4.4), (3.4, 4.5), (2.9, 5.3), (2.0, 5.2), mys]
    add_polyline(s, road, MUTED, width=5)
    add_polyline(s, [blr, mys], RED, width=3).line.dash_style = MSO_LINE_DASH_STYLE.DASH
    add_dot(s, *blr, 0.5, NAVY)
    add_dot(s, *mys, 0.5, NAVY)
    add_text(s, 5.95, 2.05, 2.0, 0.5, "Bengaluru", size=20, bold=True)
    add_text(s, 0.6, 6.0, 2.0, 0.5, "Mysuru", size=20, bold=True)
    add_badge(s, 1.6, 3.4, "about 130 km", RED, size=14, w=1.7)
    add_badge(s, 3.6, 5.0, "about 145 km by road", MUTED, size=14, w=2.6)
    add_text(s, 7.8, 1.6, 4.9, 2.4,
             ["The red line is the straight line between the two cities",
              "The grey line is the road, which bends around hills, lakes and towns"],
             size=20, bullets=True)
    word_of_the_day(s, 7.8, 4.3, 4.9, 2.0, "Straight-line distance",
                    "the shortest distance between two places, as a bird would fly.")
    add_notes(s, "A crow does not need roads, so it can fly straight. The straight line from "
                 "Bengaluru to Mysuru is about 130 km. The road is about 145 km. In Exercise 2 "
                 "the Measure tool gives you the straight-line distance. Ask: why might a "
                 "train trip be longer still?")


def add_buffer(slide, cx, cy, r, tint, colour):
    """A shaded circle with a dashed edge: everything within r of the centre."""
    zone = add_shape(slide, MSO_SHAPE.OVAL, cx - r, cy - r, 2 * r, 2 * r, tint, line=colour)
    zone.line.width = Pt(2.5)
    zone.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    return zone


def slide_buffers(prs):
    s = new_slide(prs)
    add_title(s, "Draw a circle around a place: buffers")
    cx, cy, r = 3.3, 4.3, 2.3
    add_buffer(s, cx, cy, r, BLUE_TINT, BLUE)
    add_arrow(s, cx, cy, cx + r - 0.05, cy, BLUE, width=3)
    add_badge(s, cx + 0.6, cy - 0.6, "500 m", BLUE, size=16, w=1.1)
    add_dot(s, cx, cy, 0.5, RED)
    add_badge(s, cx - 0.85, cy + 0.4, "Taj Mahal", NAVY, size=14, w=1.5)
    add_text(s, 6.6, 1.5, 6.1, 1.4,
             "In GeoLibre, click a place and open Quick analysis. It can draw a 500 m "
             "buffer: a circle that covers every spot within 500 m of the place.", size=20)
    word_of_the_day(s, 6.6, 3.05, 6.1, 1.5, "Buffer",
                    "a zone that covers everything within a set distance of a place.")
    add_card(s, 6.6, 4.8, 6.1, 1.9, "How planners use it",
             "No new buildings are allowed within 100 m of a protected monument like the "
             "Taj Mahal. For 200 m beyond that, building needs special permission.",
             GREEN_TINT, GREEN, size=17)
    add_notes(s, "A buffer is like drawing a circle with a compass, with the place at the "
                 "centre. Everything inside the circle is within that distance. India's law "
                 "for ancient monuments (the AMASR Act) uses two buffers around every protected "
                 "monument: a 100 m prohibited area where nothing new can be built, and a "
                 "further 200 m regulated area where building needs permission from the "
                 "National Monuments Authority. Planners draw these buffers in GIS to check "
                 "whether a new building is allowed. You will make your own buffers in "
                 "Exercise 3.")


def slide_metro(prs):
    s = new_slide(prs)
    add_title(s, '"500 m from the metro": what does it mean?')
    add_picture(s, IMAGES / "bengaluru_metro.jpg", 0.6, 1.3, h=5.6)
    add_caption(s, 0.6, 6.95, 6.3, "Map by Sonia Das: metro stations in Bengaluru (2024)")
    add_text(s, 7.3, 1.35, 5.4, 0.5, "Zoom in on one station", size=22, bold=True,
             color=NAVY)
    cx, cy, r = 9.9, 3.75, 1.55
    add_buffer(s, cx, cy, r, GREEN_TINT, GREEN)
    homes = [(9.0, 3.1, True), (10.7, 4.5, True), (10.5, 2.75, True),
             (7.75, 2.3, False), (12.0, 3.2, False), (8.1, 5.3, False)]
    for hx, hy, inside in homes:
        add_shape(s, MSO_SHAPE.RECTANGLE, hx - 0.15, hy - 0.15, 0.3, 0.3,
                  GREEN if inside else MUTED, line=WHITE)
    add_dot(s, cx, cy, 0.45, RGBColor(0x5B, 0x3F, 0xA8))  # purple, like the Purple Line
    add_badge(s, cx - 0.6, cy + 0.35, "Metro", NAVY, size=14, w=1.2)
    add_shape(s, MSO_SHAPE.RECTANGLE, 7.4, 5.65, 0.3, 0.3, GREEN, line=WHITE)
    add_text(s, 7.8, 5.55, 4.9, 0.5, "Inside: within 500 m of the station", size=16)
    add_shape(s, MSO_SHAPE.RECTANGLE, 7.4, 6.15, 0.3, 0.3, MUTED, line=WHITE)
    add_text(s, 7.8, 6.05, 4.9, 0.5, "Outside: further than 500 m", size=16)
    add_badge(s, 7.4, 6.6, "500 m is about a 6 minute walk", BLUE, size=14, w=3.6)
    add_notes(s, "This is a real map of Bengaluru's Purple and Green metro lines. On a city map "
                 "a 500 m circle is tiny, so on the right we zoom in on one station. When an "
                 "advert says a flat is 500 m from the metro, it means the flat is inside a "
                 "500 m buffer around the station: about a 6 minute walk. Bengaluru's "
                 "Transit-Oriented Development policy (2022) uses this same 500 m zone to plan "
                 "more homes, offices and footpaths near stations, so more people can walk to "
                 "the metro. Ask: is your home or school near a station? Find Indiranagar or "
                 "Rajajinagar on the map.")


def slide_gis_around_us(prs):
    s = new_slide(prs)
    add_title(s, "GIS is all around us")
    amber_label = RGBColor(0x92, 0x40, 0x0E)
    cards = [("Home adverts", '"Flat for sale, 500 m from the metro." That is a buffer '
              "around a station.", BLUE_TINT, BLUE),
             ("City planning", "Bengaluru plans more homes and offices within 500 m of "
              "metro stations.", GREEN_TINT, GREEN),
             ("Monuments", "No new buildings within 100 m of a protected monument.",
              AMBER_TINT, amber_label),
             ("Schools", "Young children should have a primary school within 1 km of home.",
              BLUE_TINT, BLUE),
             ("Food delivery", "Apps only show restaurants close enough to deliver to you.",
              GREEN_TINT, GREEN),
             ("Emergencies", "An ambulance is sent from the nearest hospital.", AMBER_TINT,
              amber_label)]
    for i, (label, body, tint, colour) in enumerate(cards):
        x = 0.6 + (i % 3) * 4.1
        y = 1.5 + (i // 3) * 2.2
        add_card(s, x, y, 3.9, 2.0, label, body, tint, colour, size=17)
    add_text(s, 0.6, 6.0, 12.1, 0.9,
             "Next time you see a distance in an advert or an app, think: that's GIS!",
             size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_notes(s, "Every one of these is a question about distance from a place, which is what a "
                 "GIS answers. The school rule comes from Karnataka's Right to Education rules: "
                 "a primary school within 1 km walking distance. Delivery apps draw a zone "
                 "around each restaurant. Emergency services look for the nearest hospital or "
                 "fire station. Ask: where else have you seen a distance like this? In "
                 "Exercise 3 you will draw these zones around your own city.")


def slide_map_checklist(prs):
    s = new_slide(prs)
    add_title(s, "Your map checklist")
    items = ["Title: what is the map about?", "North arrow: which way is north?",
             "Scale bar: how far is it really?", "Legend: what do the colours mean?"]
    y = 1.7
    for i, item in enumerate(items, start=1):
        number_dot(s, 0.95, y + 0.3, str(i), fill=GREEN if i < 4 else AMBER)
        add_text(s, 1.45, y, 6.0, 0.6, item, size=22, bold=i == 4)
        y += 1.0
    add_text(s, 0.6, 5.85, 6.8, 0.8, "You added 1 to 3 in Exercise 1. The legend is new.",
             size=18, color=MUTED)
    # a sample legend for the four types of place
    box = add_shape(s, MSO_SHAPE.RECTANGLE, 8.3, 1.7, 4.0, 3.9, WHITE, line=TEXT)
    box.line.width = Pt(1.5)
    add_text(s, 8.55, 1.85, 3.5, 0.5, "Legend", size=22, bold=True)
    y = 2.6
    for name, colour in TYPE_COLOURS:
        add_dot(s, 8.85, y + 0.25, 0.4, colour)
        add_text(s, 9.3, y, 2.8, 0.5, name, size=20)
        y += 0.7
    add_notes(s, "Every good map has these four parts. Exercise 1 had a title, north arrow and "
                 "scale bar. Because our dots now have different colours, the reader needs a "
                 "legend to know what they mean. Without it, a red dot could mean anything.")


def slide_quiz(prs):
    s = new_slide(prs, dark=True)
    add_text(s, 0.6, 0.4, 12.1, 0.9, "Quick quiz", size=36, bold=True, color=WHITE)
    questions = ["Which tells you how far north or south: latitude or longitude?",
                 "Is India north or south of the equator?",
                 "What does CSV stand for?",
                 "Why is a road trip longer than the straight-line distance?",
                 "What does \"500 m from the metro\" mean?"]
    y = 1.5
    for i, q in enumerate(questions, start=1):
        add_dot(s, 0.95, y + 0.28, 0.5, AMBER, outline=AMBER)
        add_text(s, 0.7, y + 0.02, 0.5, 0.5, str(i), size=20, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER)
        add_text(s, 1.45, y, 7.0, 0.9, q, size=20, color=WHITE)
        y += 0.95
    card = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 8.8, 1.6, 3.9, 3.9, GREEN)
    card.adjustments[0] = 0.08
    add_text(s, 9.1, 1.85, 3.3, 0.5, "NEXT UP", size=14, bold=True, color=AMBER_TINT)
    add_text(s, 9.1, 2.35, 3.3, 1.6, "Exercise 2: Famous Places of India", size=28,
             bold=True, color=WHITE)
    add_text(s, 9.1, 4.0, 3.3, 1.4, "Turn a table of 16 places into your own map.",
             size=18, color=WHITE)
    add_notes(s, "Answers: 1. Latitude. 2. North, all of it. 3. Comma-Separated Values. "
                 "4. Roads bend around hills, lakes and buildings, but a straight line does "
                 "not. 5. The place is inside a 500 m buffer around a metro station, about a "
                 "6 minute walk. Then hand out the Exercise 2 tutorial.")


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for build in [slide_title, slide_journey, slide_two_addresses, slide_latitude,
                  slide_longitude, slide_india_grid, slide_find_the_place, slide_table,
                  slide_rows_to_dots, slide_sorting, slide_maps_know_facts, slide_crow_flies,
                  slide_buffers, slide_metro, slide_gis_around_us, slide_map_checklist, slide_quiz]:
        build(prs)
    prs.save(OUTPUT)
    print(f"Saved {OUTPUT} with {len(prs.slides)} slides")


if __name__ == "__main__":
    main()
