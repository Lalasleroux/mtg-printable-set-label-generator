import argparse
import base64
import logging
import os
from datetime import datetime
from pathlib import Path

import jinja2
import requests


log = logging.getLogger(__name__)

BASE_DIR = Path(os.path.abspath(os.path.dirname(__file__)))
CAIRO_DLL_DIR = os.environ.get(
    "CAIRO_DLL_DIR", r"C:\\Program Files\\GTK3-Runtime Win64\\bin"
)

ENV = jinja2.Environment(
    loader=jinja2.FileSystemLoader(BASE_DIR / "templates"),
    autoescape=jinja2.select_autoescape(["html", "xml"]),
)

def _prepare_cairo():
    """Attempt to add Cairo DLL path (Windows) then import cairosvg.

    If unsuccessful, returns None and PDF generation will be skipped.
    """
    # Add DLL directory only if it exists to avoid FileNotFoundError
    if hasattr(os, "add_dll_directory") and CAIRO_DLL_DIR and os.path.isdir(CAIRO_DLL_DIR):
        try:
            os.add_dll_directory(CAIRO_DLL_DIR)
        except Exception:  # pragma: no cover - defensive
            pass
    try:
        import cairosvg  # type: ignore
        return cairosvg
    except Exception as e:  # pragma: no cover - environment dependent
        # Defer logging until logging configured; use print as fallback
        try:
            log.warning(
                "CairoSVG indisponible (%s). Les PDFs ne seront pas générés, seulement les SVG.",
                e,
            )
        except Exception:
            print("[WARN] CairoSVG not available, PDFs will be skipped:", e)
        return None

# Set types we are interested in
SET_TYPES = (
    "core",
    "expansion",
    "starter",  # Portal, P3k, welcome decks
    "masters",
    "commander",
    "planechase",
    "draft_innovation",  # Battlebond, Conspiracy
    "duel_deck",  # Duel Deck Elves,
    "premium_deck",  # Premium Deck Series: Slivers, Premium Deck Series: Graveborn
    "from_the_vault",  # Make sure to adjust the MINIMUM_SET_SIZE if you want these
    "archenemy",
    "box",
    "funny",  # Unglued, Unhinged, Ponies: TG, etc.
    # "memorabilia",  # Commander's Arsenal, Celebration Cards, World Champ Decks
    # "spellbook",
    # These are relatively large groups of sets
    # You almost certainly don't want these
    # "token",
    # "promo",
)

# Only include sets at least this size
# For reference, the smallest proper expansion is Arabian Nights with 78 cards
MINIMUM_SET_SIZE = 50

# Set codes you might want to ignore
IGNORED_SETS = (
    "cmb1",  # Mystery Booster Playtest Cards
    "amh1",  # Modern Horizon Art Series
    "cmb2",  # Mystery Booster Playtest Cards Part Deux
)

# Used to rename very long set names
RENAME_SETS = {
    "Adventures in the Forgotten Realms": "Forgotten Realms",
    "Adventures in the Forgotten Realms Minigames": "Forgotten Realms Minigames",
    "Angels: They're Just Like Us but Cooler and with Wings": "Angels: Just Like Us",
    "Archenemy: Nicol Bolas Schemes": "Archenemy: Bolas Schemes",
    "Chronicles Foreign Black Border": "Chronicles FBB",
    "Commander Anthology Volume II": "Commander Anthology II",
    "Commander Legends: Battle for Baldur's Gate": "CMDR Legends: Baldur's Gate",
    "Dominaria United Commander": "Dominaria United [C]",
    "Duel Decks: Elves vs. Goblins": "DD: Elves vs. Goblins",
    "Duel Decks: Jace vs. Chandra": "DD: Jace vs. Chandra",
    "Duel Decks: Divine vs. Demonic": "DD: Divine vs. Demonic",
    "Duel Decks: Garruk vs. Liliana": "DD: Garruk vs. Liliana",
    "Duel Decks: Phyrexia vs. the Coalition": "DD: Phyrexia vs. Coalition",
    "Duel Decks: Elspeth vs. Tezzeret": "DD: Elspeth vs. Tezzeret",
    "Duel Decks: Knights vs. Dragons": "DD: Knights vs. Dragons",
    "Duel Decks: Ajani vs. Nicol Bolas": "DD: Ajani vs. Nicol Bolas",
    "Duel Decks: Heroes vs. Monsters": "DD: Heroes vs. Monsters",
    "Duel Decks: Speed vs. Cunning": "DD: Speed vs. Cunning",
    "Duel Decks Anthology: Elves vs. Goblins": "DDA: Elves vs. Goblins",
    "Duel Decks Anthology: Jace vs. Chandra": "DDA: Jace vs. Chandra",
    "Duel Decks Anthology: Divine vs. Demonic": "DDA: Divine vs. Demonic",
    "Duel Decks Anthology: Garruk vs. Liliana": "DDA: Garruk vs. Liliana",
    "Duel Decks: Elspeth vs. Kiora": "DD: Elspeth vs. Kiora",
    "Duel Decks: Zendikar vs. Eldrazi": "DD: Zendikar vs. Eldrazi",
    "Duel Decks: Blessed vs. Cursed": "DD: Blessed vs. Cursed",
    "Duel Decks: Nissa vs. Ob Nixilis": "DD: Nissa vs. Ob Nixilis",
    "Duel Decks: Merfolk vs. Goblins": "DD: Merfolk vs. Goblins",
    "Duel Decks: Elves vs. Inventors": "DD: Elves vs. Inventors",
    "Duel Decks: Mirrodin Pure vs. New Phyrexia": "DD: Mirrodin vs.New Phyrexia",
    "Duel Decks: Izzet vs. Golgari": "Duel Decks: Izzet vs. Golgari",
    "Fourth Edition Foreign Black Border": "Fourth Edition FBB",
    "Global Series Jiang Yanggu & Mu Yanling": "Jiang Yanggu & Mu Yanling",
    "Innistrad: Crimson Vow Minigames": "Crimson Vow Minigames",
    "Introductory Two-Player Set": "Intro Two-Player Set",
    "March of the Machine: The Aftermath": "MotM: The Aftermath",
    "March of the Machine Commander": "March of the Machine [C]",
    "Murders at Karlov Manor Commander": "Murders at Karlov Manor [C]",
    "Mystery Booster Playtest Cards": "Mystery Booster Playtest",
    "Mystery Booster Playtest Cards 2019": "MB Playtest Cards 2019",
    "Mystery Booster Playtest Cards 2021": "MB Playtest Cards 2021",
    "Mystery Booster Retail Edition Foils": "Mystery Booster Retail Foils",
    "Outlaws of Thunder Junction Commander": "Outlaws of Thunder Junction [C]",
    "Phyrexia: All Will Be One Commander": "Phyrexia: All Will Be One [C]",
    "Planechase Anthology Planes": "Planechase Anth. Planes",
    "Premium Deck Series: Slivers": "Premium Deck Slivers",
    "Premium Deck Series: Graveborn": "Premium Deck Graveborn",
    "Premium Deck Series: Fire and Lightning": "PD: Fire & Lightning",
    "Shadows over Innistrad Remastered": "SOI Remastered",
    "Strixhaven: School of Mages Minigames": "Strixhaven Minigames",
    "Tales of Middle-earth Commander": "Tales of Middle-earth [C]",
    "The Brothers' War Retro Artifacts": "Brothers' War Retro",
    "The Brothers' War Commander": "Brothers' War Commander",
    "The Lord of the Rings: Tales of Middle-earth": "LOTR: Tales of Middle-earth",
    "The Lost Caverns of Ixalan Commander": "The Lost Caverns of Ixalan [C]",
    "Warhammer 40,000 Commander": "Warhammer 40K [C]",
    "Wilds of Eldraine Commander": "Wilds of Eldraine [C]",
    "World Championship Decks 1997": "World Championship 1997",
    "World Championship Decks 1998": "World Championship 1998",
    "World Championship Decks 1999": "World Championship 1999",
    "World Championship Decks 2000": "World Championship 2000",
    "World Championship Decks 2001": "World Championship 2001",
    "World Championship Decks 2002": "World Championship 2002",
    "World Championship Decks 2003": "World Championship 2003",
    "World Championship Decks 2004": "World Championship 2004",
}


class LabelGenerator:

    DEFAULT_OUTPUT_DIR = Path(os.getcwd()) / "output"

    COLS = 4
    ROWS = 15
    MARGIN = 100  # in 1/10 mm
    START_X = MARGIN
    START_Y = MARGIN

    PAPER_SIZES = {
        "letter": {"width": 2790, "height": 2160, },  # in 1/10 mm
        "a4": {"width": 2970, "height": 2100, },
    }
    DEFAULT_PAPER_SIZE = "letter"

    def __init__(self, paper_size=DEFAULT_PAPER_SIZE, output_dir=DEFAULT_OUTPUT_DIR, generate_pdfs=True, include_color_labels=False, portrait=False, margin_mm=None):
        """Initialize a label generator with layout and feature flags."""
        # Store config
        self.paper_size = paper_size
        self.portrait = portrait
        paper = self.PAPER_SIZES[paper_size].copy()
        # Swap dimensions if portrait requested
        if portrait:
            paper["width"], paper["height"] = paper["height"], paper["width"]

        # Filters / selection parameters
        self.set_codes = []  # optionally narrowed later by CLI args
        self.ignored_sets = IGNORED_SETS
        self.set_types = SET_TYPES
        self.minimum_set_size = MINIMUM_SET_SIZE

        # Paper geometry & margin
        self.width = paper["width"]
        self.height = paper["height"]
        if margin_mm is not None:
            # Convert mm to internal units (1/10 mm)
            self.MARGIN = int(round(margin_mm * 10))
        # Recompute derived spacing using (possibly overridden) margin
        self.delta_x = (self.width - (2 * self.MARGIN)) / self.COLS
        self.delta_y = (self.height - (2 * self.MARGIN)) / self.ROWS

        # Output / feature flags
        self.output_dir = Path(output_dir)
        self.generate_pdfs = generate_pdfs
        self.include_color_labels = include_color_labels

    def generate_labels(self, sets=None):
        if sets:
            self.ignored_sets = ()
            self.minimum_set_size = 0
            self.set_types = ()
            self.set_codes = [exp.lower() for exp in sets]
        cairosvg_mod = _prepare_cairo() if self.generate_pdfs else None

        raw_items = []
        if self.include_color_labels:
            raw_items.extend(self.get_color_label_raw())
        # existing set labels
        raw_items.extend(self.get_set_label_raw())

        labels = self.layout_labels(raw_items)
        page = 1
        while labels:
            exps = []
            while labels and len(exps) < (self.ROWS * self.COLS):
                exps.append(labels.pop(0))

            # Render the label template
            template = ENV.get_template("labels.svg")
            output = template.render(
                labels=exps,
                horizontal_guides=self.create_horizontal_cutting_guides(),
                vertical_guides=self.create_vertical_cutting_guides(),
                WIDTH=self.width,
                HEIGHT=self.height,
            )
            orient = "-portrait" if self.portrait else ""
            outfile_svg = self.output_dir / f"labels-{self.paper_size}{orient}-{page:02}.svg"
            outfile_pdf = str(
                self.output_dir / f"labels-{self.paper_size}{orient}-{page:02}.pdf"
            )

            log.info(f"Writing {outfile_svg}...")
            with open(outfile_svg, "w") as fd:
                fd.write(output)

            if cairosvg_mod and self.generate_pdfs:
                log.info(f"Writing {outfile_pdf}...")
                with open(outfile_svg, "rb") as fd:
                    cairosvg_mod.svg2pdf(file_obj=fd, write_to=outfile_pdf)
            else:
                log.info(
                    "CairoSVG absent: PDF non généré pour la page %s (SVG disponible)",
                    page,
                )

            page += 1

    def get_set_data(self):
        log.info("Getting set data and icons from Scryfall")

        # https://scryfall.com/docs/api/sets
        # https://scryfall.com/docs/api/sets/all
        resp = requests.get("https://api.scryfall.com/sets")
        resp.raise_for_status()

        data = resp.json()["data"]
        set_data = []
        for exp in data:
            if exp["code"] in self.ignored_sets:
                continue
            elif exp["card_count"] < self.minimum_set_size:
                continue
            elif self.set_types and exp["set_type"] not in self.set_types:
                continue
            elif self.set_codes and exp["code"].lower() not in self.set_codes:
                # Scryfall set codes are always lowercase
                continue
            else:
                set_data.append(exp)

        # Warn on any unknown set codes
        if self.set_codes:
            known_sets = set([exp["code"] for exp in data])
            specified_sets = set([code.lower() for code in self.set_codes])
            unknown_sets = specified_sets.difference(known_sets)
            for set_code in unknown_sets:
                log.warning("Unknown set '%s'", set_code)

        set_data.reverse()
        return set_data

    def get_set_label_raw(self):
        raw = []
        for exp in self.get_set_data():
            name = RENAME_SETS.get(exp["name"], exp["name"])
            icon_resp = requests.get(exp["icon_svg_uri"])
            icon_b64 = None
            if icon_resp.ok:
                icon_b64 = base64.b64encode(icon_resp.content).decode('utf-8')
            raw.append(
                {
                    "name": name,
                    "code": exp["code"],
                    "date": datetime.strptime(exp["released_at"], "%Y-%m-%d").date(),
                    "icon_url": exp["icon_svg_uri"],
                    "icon_b64": icon_b64,
                }
            )
        return raw

    COLOR_SYMBOLS = [
        ("W", "White"),
        ("U", "Blue"),
        ("B", "Black"),
        ("R", "Red"),
        ("G", "Green"),
        ("C", "Colorless"),
    ]

    def get_color_label_raw(self):
        try:
            resp = requests.get("https://api.scryfall.com/symbology")
            resp.raise_for_status()
            data = resp.json().get("data", [])
            symbol_map = {item.get("symbol").strip("{}" ): item for item in data if item.get("svg_uri")}
        except Exception as e:
            log.warning("Failed fetching symbology: %s", e)
            return []
        raw = []
        for code, name in self.COLOR_SYMBOLS:
            info = symbol_map.get(code)
            if not info:
                continue
            icon_resp = requests.get(info["svg_uri"])
            icon_b64 = None
            if icon_resp.ok:
                icon_b64 = base64.b64encode(icon_resp.content).decode('utf-8')
            raw.append(
                {
                    "name": name,
                    "code": code,
                    "date": None,
                    "icon_url": info["svg_uri"],
                    "icon_b64": icon_b64,
                }
            )
        return raw

    def layout_labels(self, raw_items):
        labels = []
        for idx, item in enumerate(raw_items):
            col = (idx // self.ROWS) % self.COLS
            row = idx % self.ROWS
            x = self.START_X + col * self.delta_x
            y = self.START_Y + row * self.delta_y
            labels.append({**item, "x": x, "y": y})
        return labels

    # --- Color + Set combined wide label mode ---
    WIDE_LABEL_MM = 88  # requested width in mm

    def generate_color_set_mode(self, sets=None):
        """Generate wide labels: one per (set, color)."""
        if sets:
            self.ignored_sets = ()
            self.minimum_set_size = 0
            self.set_types = ()
            self.set_codes = [exp.lower() for exp in sets]

        # Fetch once
        color_entries = self.get_color_label_raw()
        color_map = {e["code"]: e for e in color_entries}
        set_data = self.get_set_data()

        # Pre-fetch set icons (avoid duplicate downloads per color)
        set_icon_cache = {}
        labels_raw = []
        for exp in set_data:
            set_name = RENAME_SETS.get(exp["name"], exp["name"])
            if exp["code"] not in set_icon_cache:
                icon_resp = requests.get(exp["icon_svg_uri"])
                icon_b64 = None
                if icon_resp.ok:
                    icon_b64 = base64.b64encode(icon_resp.content).decode('utf-8')
                set_icon_cache[exp["code"]] = icon_b64
            set_icon_b64 = set_icon_cache[exp["code"]]
            for code, _cname in self.COLOR_SYMBOLS:
                cinfo = color_map.get(code)
                if not cinfo:
                    continue
                labels_raw.append({
                    "set_name": set_name,
                    "set_code": exp["code"],
                    "set_icon_b64": set_icon_b64,
                    "color_code": code,
                    "color_name": cinfo["name"],
                    "color_icon_b64": cinfo["icon_b64"],
                })

        laid_out = self.layout_wide_labels(labels_raw)

        # Render pages
        template = ENV.get_template("labels_color_sets.svg")
        page = 1
        cairosvg_mod = _prepare_cairo() if self.generate_pdfs else None
        # Page capacity must use the dynamically computed wide column count, not the default COLS
        per_page = getattr(self, "_wide_cols", self.COLS) * self.ROWS
        while laid_out:
            batch = []
            while laid_out and len(batch) < per_page:
                batch.append(laid_out.pop(0))
            output = template.render(
                labels=batch,
                horizontal_guides=self.create_horizontal_cutting_guides(),
                vertical_guides=self._wide_vertical_guides(),
                WIDTH=self.width,
                HEIGHT=self.height,
            )
            orient = "-portrait" if self.portrait else ""
            outfile_svg = self.output_dir / f"labels-color-set-{self.paper_size}{orient}-{page:02}.svg"
            outfile_pdf = str(self.output_dir / f"labels-color-set-{self.paper_size}{orient}-{page:02}.pdf")
            log.info(f"Writing {outfile_svg}...")
            with open(outfile_svg, "w") as fd:
                fd.write(output)
            if cairosvg_mod and self.generate_pdfs:
                log.info(f"Writing {outfile_pdf}...")
                with open(outfile_svg, "rb") as fd:
                    cairosvg_mod.svg2pdf(file_obj=fd, write_to=outfile_pdf)
            page += 1

    def layout_wide_labels(self, raw_items):
        """Layout wide labels using fixed physical width (WIDE_LABEL_MM)."""
        cell_w = self.WIDE_LABEL_MM * 10  # convert mm to internal units
        usable = self.width - 2 * self.MARGIN
        cols = max(1, int(usable // cell_w))
        self._wide_cols = cols  # store for guides
        labels = []
        for idx, item in enumerate(raw_items):
            col = idx % cols
            row = (idx // cols) % self.ROWS
            page_index = idx // (cols * self.ROWS)
            x = self.START_X + col * cell_w
            y = self.START_Y + row * self.delta_y
            labels.append({**item, "x": x, "y": y, "width": cell_w, "_page": page_index})
        return labels

    def _wide_vertical_guides(self):
        if not hasattr(self, "_wide_cols"):
            return []
        cell_w = self.WIDE_LABEL_MM * 10
        guides = []
        for i in range(self._wide_cols + 1):
            x = self.MARGIN + i * cell_w
            if x > (self.width - self.MARGIN):
                break
            guides.append({
                "x1": x,
                "x2": x,
                "y1": self.MARGIN / 2,
                "y2": self.MARGIN * 0.8,
            })
            guides.append({
                "x1": x,
                "x2": x,
                "y1": self.height - self.MARGIN / 2,
                "y2": self.height - self.MARGIN * 0.8,
            })
        return guides

    def create_horizontal_cutting_guides(self):
        """Create horizontal cutting guides to help cut the labels out straight"""
        horizontal_guides = []
        for i in range(self.ROWS + 1):
            horizontal_guides.append(
                {
                    "x1": self.MARGIN / 2,
                    "x2": self.MARGIN * 0.8,
                    "y1": self.MARGIN + i * self.delta_y,
                    "y2": self.MARGIN + i * self.delta_y,
                }
            )
            horizontal_guides.append(
                {
                    "x1": self.width - self.MARGIN / 2,
                    "x2": self.width - self.MARGIN * 0.8,
                    "y1": self.MARGIN + i * self.delta_y,
                    "y2": self.MARGIN + i * self.delta_y,
                }
            )

        return horizontal_guides

    def create_vertical_cutting_guides(self):
        """Create horizontal cutting guides to help cut the labels out straight"""
        vertical_guides = []
        for i in range(self.COLS + 1):
            vertical_guides.append(
                {
                    "x1": self.MARGIN + i * self.delta_x,
                    "x2": self.MARGIN + i * self.delta_x,
                    "y1": self.MARGIN / 2,
                    "y2": self.MARGIN * 0.8,
                }
            )
            vertical_guides.append(
                {
                    "x1": self.MARGIN + i * self.delta_x,
                    "x2": self.MARGIN + i * self.delta_x,
                    "y1": self.height - self.MARGIN / 2,
                    "y2": self.height - self.MARGIN * 0.8,
                }
            )

        return vertical_guides


def main():
    log_format = '[%(levelname)s] %(message)s'
    logging.basicConfig(format=log_format, level=logging.INFO)

    parser = argparse.ArgumentParser(description="Generate MTG labels")

    parser.add_argument(
        "--output-dir",
        default=LabelGenerator.DEFAULT_OUTPUT_DIR,
        help="Output labels to this directory",
    )
    parser.add_argument(
        "--paper-size",
        default=LabelGenerator.DEFAULT_PAPER_SIZE,
        choices=LabelGenerator.PAPER_SIZES.keys(),
        help='Use this paper size (default: "letter")',
    )
    parser.add_argument(
        "sets",
        nargs="*",
        help=(
            "Only output sets with the specified set code (eg. MH1, NEO). "
            "This can be used multiple times."
        ),
        metavar="SET",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Do not generate PDF files (only SVG). Useful if Cairo / libcairo is not installed.",
    )
    parser.add_argument(
        "--color-labels",
        action="store_true",
        help="Add a label for each MTG color (W,U,B,R,G,C).",
    )
    parser.add_argument(
        "--color-set-mode",
        action="store_true",
        help="Generate wide 88mm labels: one per (set,color) with color icon left, set name centered, set icon right.",
    )
    parser.add_argument(
        "--portrait",
        action="store_true",
        help="Use portrait orientation (swap paper width/height). Useful with --color-set-mode to get exactly two 88mm columns on Letter/A4.",
    )
    parser.add_argument(
        "--commands-help",
        action="store_true",
        help="Show a concise English help about each custom option and exit.",
    )
    parser.add_argument(
        "--margin-mm",
        type=float,
        help="Override default margin (20mm) with a custom value in millimeters.",
    )

    args = parser.parse_args()

    if args.commands_help:
        print("Custom options (English):")
        print("  --paper-size {letter|a4} : Select paper size (default letter).")
        print("  --no-pdf                 : Skip PDF generation (only SVG output).")
        print("  --color-labels           : Prepend one label per MTG color (W,U,B,R,G,C).")
        print("  --color-set-mode         : Wide 88mm labels, one per (set,color) combination.")
        print("  --portrait               : Portrait orientation (swaps width/height). Useful with wide mode.")
        print("  --commands-help          : Show this custom help list and exit.")
        print("  --margin-mm N            : Set page margin to N millimeters (overrides default 20mm).")
        print("  SET codes (positional)   : Limit generation to specified set codes (e.g. lea mh1).")
        print("Examples:")
        print("  python mtglabels/generator.py --color-labels")
        print("  python mtglabels/generator.py --color-set-mode --portrait lea")
        print("  python mtglabels/generator.py --paper-size=a4 --no-pdf")
        return

    generator = LabelGenerator(
        args.paper_size,
        args.output_dir,
        generate_pdfs=not args.no_pdf,
        include_color_labels=args.color_labels,
        portrait=args.portrait,
        margin_mm=args.margin_mm,
    )
    if args.color_set_mode:
        generator.generate_color_set_mode(args.sets)
    else:
        generator.generate_labels(args.sets)


if __name__ == "__main__":
    main()
