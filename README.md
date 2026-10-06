# Mapping with Kids

A small GIS course for kids, built around [GeoLibre](https://geolibre.app), a
free and open-source web GIS tool.

**Live site:** https://soniadas123.github.io/mapping-with-kids/

## What's here

- `index.html` — the cover page, linking to the PDFs below.
- `Remote_Sensing_and_GIS_Intro.pdf` — a short intro to remote sensing,
  satellites, and GIS, meant to be read before Exercise 1.
- `exercise1/` — Exercise 1: building a map of India.
  - `Intro_to_Mapping.pptx` — a 17-slide intro deck for the teacher to present
    before the exercise: remote sensing, satellites, land cover change in
    Bengaluru (IISc data), GIS layers, points, lines
    and polygons, and the parts of a map. Speaker notes are on every slide.
  - `build_intro_pptx.py` — builds the deck with python-pptx. Edit it, then
    run `python build_intro_pptx.py` from the `exercise1/` folder in the
    `mapping_kids` conda env.
  - `images/` — photos used in the deck, cropped from the 2021
    "Mapping for Kids" blog slides.
  - `Map_of_India_Tutorial.pdf` — the step-by-step tutorial.
  - `data/` — the GeoJSON boundary data used in the exercise (the large
    `.geojson` files are not tracked in this repo; see below).
  - `screenshot/` — reference screenshots from the exercise.
- `exercise2/` — Exercise 2: putting famous places of India on the map.
  - `Intro_to_Coordinates.pptx` — a 17-slide intro deck for the teacher to
    present before the exercise: latitude and longitude, India on the grid,
    tables as point data, categorized colours, labels, straight-line distance,
    buffers (what "500 m from the metro" means, using a map of Bengaluru's
    metro), everyday uses of GIS, and the legend. Speaker notes are on every slide.
  - `build_intro_pptx.py` — builds the deck with python-pptx, reusing the
    helpers from `exercise1/build_intro_pptx.py`. Run `python make_intro_maps.py`
    once (it makes the India maps from the Exercise 1 outline, and a smaller
    copy of `images/editmetrofina_final_placelabels.png`), then
    `python build_intro_pptx.py`, from the `exercise2/` folder in the
    `mapping_kids` conda env.
  - `Famous_Places_Tutorial.pdf` — the step-by-step tutorial.
  - `Famous_Places_Tutorial.html` — the source for the PDF. Edit this,
    then rebuild the PDF from the `exercise2/` folder with Edge:
    `msedge --headless --print-to-pdf=Famous_Places_Tutorial.pdf Famous_Places_Tutorial.html`
  - `screenshot/s2.png` — example image used on the cover. The steps have
    no screenshots on purpose, so students build their own map.
  - `data/famous_places.csv` — 16 landmarks with latitude and longitude.
    The India outline is reused from Exercise 1.
- `exercise3/` — Exercise 3: drawing distance zones (buffers) around your city.
  - `Near_or_Far_Tutorial.pdf` — the step-by-step tutorial.
  - `Near_or_Far_Tutorial.html` — the source for the PDF. Rebuild it from the
    `exercise3/` folder with Edge:
    `msedge --headless --print-to-pdf=Near_or_Far_Tutorial.pdf Near_or_Far_Tutorial.html`
  - `data/my_city.csv` — a one-row starting point (Bangalore) that students
    change to their own city. Reuses the Exercise 1 and 2 data.
- `exercise4/` — Exercise 4: exploring the Himalayas with 3D terrain.
  - `Himalayas_3D_Tutorial.pdf` — the step-by-step tutorial.
  - `Himalayas_3D_Tutorial.html` — the source for the PDF. Rebuild it from the
    `exercise4/` folder with Edge:
    `msedge --headless --print-to-pdf=Himalayas_3D_Tutorial.pdf Himalayas_3D_Tutorial.html`
  - No new data: GeoLibre loads elevation tiles itself, and the India outline
    is reused from Exercise 1.

## Note on large files

A few large working files are intentionally excluded from this repo (see
`.gitignore`) since they aren't needed to view the tutorials:
`exercise1/screenshot/exercise1.html`, `exercise1/screenshot/exercise1.geolibre.json`,
and the raw `.geojson` boundary data in `exercise1/data/`. Those stay on the
original working machine.
