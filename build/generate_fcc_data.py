"""
Generate a synthetic, FCC-BDC-shaped broadband availability extract for the
Case Study 1 (whitespace analysis for "Lakeshore Broadband", an imaginary
regional operator in northern Illinois).

The schema mirrors the FCC Broadband Data Collection fixed-availability
download (frn, provider_id, brand_name, location_id, technology,
max_advertised_download_speed, max_advertised_upload_speed, low_latency,
business_residential_code, state_usps, block_geoid, h3_res8_id).

Everything here is SYNTHETIC. County names/FIPS are real Illinois codes so the
join patterns are realistic; every provider, location and speed is invented.

Usage:
    python generate_fcc_data.py --scale 1      # ~60k locations, ~200k rows (classroom)
    python generate_fcc_data.py --scale 20     # ~1.2M locations -> breaks Excel's row limit
"""
import argparse
import numpy as np
import pandas as pd
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "data"

# county_fips: (name, type, n_locations_base, households, median_income, footprint_share, region_role)
COUNTIES = {
    "097": ("Lake",      "suburban", 9000, 265000, 98000, 0.70, "core"),
    "111": ("McHenry",   "exurban",  5500, 115000, 92000, 0.55, "core"),
    "089": ("Kane",      "suburban", 6500, 190000, 88000, 0.30, "core"),
    "007": ("Boone",     "rural",    1300,  20000, 71000, 0.10, "core"),
    "031": ("Cook",      "urban",   14000, 2050000, 78000, 0.00, "adjacent"),
    "043": ("DuPage",    "suburban", 7000, 350000, 105000, 0.00, "adjacent"),
    "037": ("DeKalb",    "rural",    1800,  38000, 68000, 0.00, "adjacent"),
    "201": ("Winnebago", "exurban",  3500, 115000, 62000, 0.00, "adjacent"),
    "093": ("Kendall",   "exurban",  2200,  45000, 110000, 0.00, "adjacent"),
    "197": ("Will",      "suburban", 6000, 240000, 100000, 0.00, "outer"),
    "099": ("LaSalle",   "rural",    2000,  45000, 66000, 0.00, "outer"),
    "091": ("Kankakee",  "rural",    1700,  42000, 64000, 0.00, "outer"),
}

# provider_id: (frn, brand_name, technology, home_types, presence_by_county_type)
PROVIDERS = {
    130001: ("0031200001", "Lakeshore Broadband", 40, None),          # our company (footprint-driven)
    130002: ("0031200002", "Prairie Fiber",       50, {"urban": .35, "suburban": .40, "exurban": .20, "rural": .05}),
    130003: ("0031200003", "MidState Cable",      40, {"urban": .85, "suburban": .60, "exurban": .35, "rural": .15}),
    130004: ("0031200004", "HeritageTel",         10, {"urban": .60, "suburban": .70, "exurban": .80, "rural": .85}),
    130005: ("0031200005", "SkyReach Wireless",   70, {"urban": .05, "suburban": .15, "exurban": .45, "rural": .70}),
    130006: ("0031200006", "OrbitNet",            60, {"urban": .95, "suburban": .97, "exurban": .98, "rural": .99}),
    130007: ("0031200007", "StarMesh",            61, {"urban": .90, "suburban": .93, "exurban": .95, "rural": .97}),
}

TECH_SPEEDS = {
    # technology: list of (down, up, weight)
    40: [(300, 20, .30), (500, 20, .25), (1000, 35, .25), (100, 10, .12), (1200, 35, .08)],
    50: [(500, 500, .35), (1000, 1000, .40), (2000, 2000, .15), (5000, 5000, .10)],
    10: [(25, 3, .30), (50, 5, .30), (100, 10, .25), (10, 1, .15)],
    70: [(50, 10, .35), (100, 20, .30), (200, 25, .20), (25, 5, .15)],
    60: [(100, 12, 1.0)],
    61: [(150, 20, .6), (100, 15, .4)],
}


def pick_speed(rng, tech, n):
    opts = TECH_SPEEDS[tech]
    w = np.array([o[2] for o in opts]); w = w / w.sum()
    idx = rng.choice(len(opts), size=n, p=w)
    down = np.array([opts[i][0] for i in idx]); up = np.array([opts[i][1] for i in idx])
    return down, up


def main(scale: float, seed: int = 7):
    rng = np.random.default_rng(seed)
    loc_rows, avail_rows = [], []
    loc_counter = 1_000_000_000
    footprint_blocks = []

    for fips, (name, ctype, n_base, hh, inc, fp_share, role) in COUNTIES.items():
        n_loc = int(n_base * scale)
        # blocks: ~35 locations per block on average
        n_blocks = max(20, n_loc // 35)
        tracts = rng.choice(np.arange(100, 990), size=max(3, n_blocks // 12), replace=False)
        block_ids = []
        for b in range(n_blocks):
            tract = tracts[b % len(tracts)]
            block = 1000 + b
            block_ids.append(f"17{fips}{tract:04d}00{block:04d}"[:15].ljust(15, "0"))
        block_ids = np.array(block_ids)

        # Lakeshore footprint at block level
        fp_mask = rng.random(n_blocks) < fp_share
        footprint_blocks += [(b, fips, name) for b in block_ids[fp_mask]]
        fp_lookup = dict(zip(block_ids, fp_mask))

        # locations
        block_of_loc = rng.choice(block_ids, size=n_loc)
        # county centroid-ish coordinates (rough Illinois lat/lon), jittered
        base_lat = {"097": 42.35, "111": 42.32, "089": 41.94, "007": 42.32, "031": 41.84, "043": 41.85,
                    "037": 41.89, "201": 42.34, "093": 41.59, "197": 41.45, "099": 41.34, "091": 41.14}[fips]
        base_lon = {"097": -87.86, "111": -88.45, "089": -88.43, "007": -88.82, "031": -87.75, "043": -88.09,
                    "037": -88.77, "201": -89.16, "093": -88.43, "197": -87.98, "099": -88.89, "091": -87.86}[fips]
        lat = base_lat + rng.normal(0, 0.12, n_loc)
        lon = base_lon + rng.normal(0, 0.15, n_loc)
        unit_count = np.where(rng.random(n_loc) < (0.08 if ctype == "urban" else 0.03),
                              rng.integers(4, 60, n_loc), 1)
        land_use = np.where(rng.random(n_loc) < 0.07, "business", "residential")
        loc_ids = np.arange(loc_counter, loc_counter + n_loc); loc_counter += n_loc

        for i in range(n_loc):
            loc_rows.append((loc_ids[i], block_of_loc[i], fips, name, "IL",
                             round(lat[i], 5), round(lon[i], 5), int(unit_count[i]), land_use[i]))

        # availability rows, provider by provider
        for pid, (frn, brand, tech, presence) in PROVIDERS.items():
            if pid == 130001:
                # Lakeshore: present in footprint blocks; ~92% of locations there
                in_fp = np.array([fp_lookup[b] for b in block_of_loc])
                mask = in_fp & (rng.random(n_loc) < 0.92)
                # 35% of footprint blocks upgraded to fiber
                fiber_blocks = set(rng.choice(block_ids[fp_mask], size=int(fp_mask.sum() * 0.35), replace=False)) if fp_mask.sum() else set()
            else:
                p = presence[ctype]
                # make presence clustered by block, not uniform per location
                block_present = dict(zip(block_ids, rng.random(n_blocks) < min(1.0, p * 1.15)))
                mask = np.array([block_present[b] for b in block_of_loc]) & (rng.random(n_loc) < 0.90)
                fiber_blocks = set()
            idx = np.where(mask)[0]
            if len(idx) == 0:
                continue
            techs = np.full(len(idx), tech)
            if pid == 130001:
                techs = np.array([50 if block_of_loc[i] in fiber_blocks else 40 for i in idx])
            if pid == 130004:  # HeritageTel has fiber in ~12% of blocks (towns)
                ft_blocks = set(rng.choice(block_ids, size=int(n_blocks * 0.12), replace=False))
                techs = np.array([50 if block_of_loc[i] in ft_blocks else 10 for i in idx])
            down = np.zeros(len(idx), dtype=int); up = np.zeros(len(idx), dtype=int)
            for t in np.unique(techs):
                m = techs == t
                d, u = pick_speed(rng, int(t), int(m.sum()))
                down[m] = d; up[m] = u
            low_lat = np.where(techs == 60, 0, 1)
            brc = rng.choice(["R", "X", "B"], size=len(idx), p=[.62, .33, .05])
            brand_out = np.full(len(idx), brand, dtype=object)
            # dirty data: ~2% inconsistent casing / whitespace in brand names
            dirty = rng.random(len(idx)) < 0.02
            brand_out[dirty] = [b.upper() if rng.random() < .5 else b + "  " for b in brand_out[dirty]]
            for j, i in enumerate(idx):
                h3 = f"88{int(abs(lat[i]) * 1e4) % 100000:05d}{int(abs(lon[i]) * 1e3) % 1000:03d}ffff"[:15]
                avail_rows.append((frn, pid, brand_out[j], int(loc_ids[i]), int(techs[j]), int(down[j]), int(up[j]),
                                   int(low_lat[j]), brc[j], "IL", block_of_loc[i], h3))

    loc_cols = ["location_id", "block_geoid", "county_fips", "county_name", "state_usps",
                "latitude", "longitude", "unit_count", "land_use"]
    av_cols = ["frn", "provider_id", "brand_name", "location_id", "technology",
               "max_advertised_download_speed", "max_advertised_upload_speed", "low_latency",
               "business_residential_code", "state_usps", "block_geoid", "h3_res8_id"]
    locations = pd.DataFrame(loc_rows, columns=loc_cols)
    avail = pd.DataFrame(avail_rows, columns=av_cols)

    # dirty data: ~1% exact duplicate rows (a real BDC quirk when providers refile)
    dups = avail.sample(frac=0.01, random_state=seed)
    avail = pd.concat([avail, dups], ignore_index=True).sample(frac=1, random_state=seed).reset_index(drop=True)

    footprint = pd.DataFrame(footprint_blocks, columns=["block_geoid", "county_fips", "county_name"]).drop_duplicates()
    demo = pd.DataFrame([(f, v[0], v[3], v[4], v[1], v[6]) for f, v in COUNTIES.items()],
                        columns=["county_fips", "county_name", "households", "median_household_income",
                                 "county_type", "relation_to_footprint"])
    tech_codes = pd.DataFrame([
        (10, "Copper Wire (DSL)"), (40, "Cable (Coaxial / HFC)"), (50, "Fiber to the Premises"),
        (60, "Geostationary Satellite"), (61, "Non-geostationary Satellite"),
        (70, "Unlicensed Fixed Wireless"), (71, "Licensed Fixed Wireless"),
        (72, "Licensed-by-Rule Fixed Wireless"), (0, "Other")], columns=["technology", "technology_name"])

    OUT.mkdir(exist_ok=True, parents=True)
    suffix = "" if scale == 1 else f"_x{int(scale)}"
    avail.to_csv(OUT / f"bdc_il_fixed_availability_sample{suffix}.csv.gz", index=False, compression="gzip")
    locations.to_csv(OUT / f"bdc_il_locations_sample{suffix}.csv.gz", index=False, compression="gzip")
    footprint.to_csv(OUT / "lakeshore_footprint_blocks.csv", index=False)
    demo.to_csv(OUT / "il_county_demographics_illustrative.csv", index=False)
    tech_codes.to_csv(OUT / "bdc_technology_codes.csv", index=False)
    print(f"availability rows: {len(avail):,} | locations: {len(locations):,} | footprint blocks: {len(footprint):,}")
    return avail, locations


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--scale", type=float, default=1.0)
    a = ap.parse_args()
    main(a.scale)
