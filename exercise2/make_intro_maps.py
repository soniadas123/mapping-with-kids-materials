"""Make the images used in the Exercise 2 intro deck.

Run inside the mapping_kids conda env:
    python make_intro_maps.py

india_grid.png    India with latitude and longitude lines, for the "find the place" game
india_places.png  India with three sample dots, to show "each row becomes a dot"
bengaluru_metro.jpg  a smaller copy of the Bengaluru metro map, for the buffer slides
"""
import json
from pathlib import Path

import geopandas as gpd
import matplotlib.pyplot as plt
from PIL import Image

HERE = Path(__file__).parent
BOUNDARY = HERE.parent / "exercise1" / "data" / "india_boundary.geojson"
IMAGES = HERE / "images"
METRO_MAP = IMAGES / "editmetrofina_final_placelabels.png"

NAVY = "#0F2340"
GREEN = "#2E7D4F"
BLUE = "#1F6FB2"
AMBER = "#E8A317"
RED = "#D9482B"
LAND = "#E6F4EA"
GRID = "#9CA3AF"


def load_india():
    """Read the Exercise 1 outline, simplify it, and switch to latitude and longitude."""
    with open(BOUNDARY, encoding="utf-8") as f:
        features = json.load(f)["features"]
    india = gpd.GeoDataFrame.from_features(features, crs="EPSG:3857")
    india["geometry"] = india.simplify(2000)  # 2 km is plenty for a slide
    return india.to_crs("EPSG:4326")


def base_map(india, figsize):
    fig, ax = plt.subplots(figsize=figsize, dpi=150)
    india.plot(ax=ax, color=LAND, edgecolor=GREEN, linewidth=1.5)
    ax.set_xlim(66, 99)
    ax.set_ylim(5, 38)
    ax.set_aspect("equal")
    for side in ax.spines.values():
        side.set_visible(False)
    return fig, ax


def make_grid_map(india):
    fig, ax = base_map(india, (7, 7))
    for lon in range(70, 100, 5):
        ax.axvline(lon, color=GRID, linewidth=0.8, zorder=0)
    for lat in range(5, 40, 5):
        ax.axhline(lat, color=GRID, linewidth=0.8, zorder=0)
    ax.axhline(23.44, color=RED, linewidth=2, linestyle="--")
    ax.text(93.8, 22.8, "Tropic of\nCancer", color=RED, fontsize=12, fontweight="bold",
            va="top")
    ax.set_xticks(range(70, 100, 5))
    ax.set_xticklabels([f"{x} E" for x in range(70, 100, 5)], fontsize=11)
    ax.set_yticks(range(5, 40, 5))
    ax.set_yticklabels([f"{y} N" for y in range(5, 40, 5)], fontsize=11)
    ax.tick_params(length=0, colors=NAVY)
    ax.set_xlabel("Longitude (east-west)", fontsize=13, color=NAVY)
    ax.set_ylabel("Latitude (north-south)", fontsize=13, color=NAVY)
    fig.savefig(IMAGES / "india_grid.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def make_places_map(india):
    fig, ax = base_map(india, (6, 6))
    ax.set_axis_off()
    places = [("Delhi", 77.21, 28.61), ("Kolkata", 88.36, 22.57), ("Bengaluru", 77.59, 12.97)]
    for name, lon, lat in places:
        ax.scatter(lon, lat, s=160, color=BLUE, edgecolor="white", linewidth=2, zorder=3)
        ax.annotate(name, (lon, lat), xytext=(10, -4), textcoords="offset points",
                    fontsize=14, fontweight="bold", color=NAVY)
    fig.savefig(IMAGES / "india_places.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def make_metro_image():
    """Save a smaller copy of the metro map, so the deck file stays small."""
    with Image.open(METRO_MAP) as image:
        width = 1600
        height = round(image.height * width / image.width)
        image.convert("RGB").resize((width, height), Image.LANCZOS).save(
            IMAGES / "bengaluru_metro.jpg", quality=85)


def main():
    IMAGES.mkdir(exist_ok=True)
    india = load_india()
    make_grid_map(india)
    make_places_map(india)
    make_metro_image()
    print(f"Saved images to {IMAGES}")


if __name__ == "__main__":
    main()
