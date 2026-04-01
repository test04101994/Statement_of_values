# """
# Test script for AWS Location Services:
#   1. Geocode an address → get lat/long
#   2. Reverse geocode lat/long → get address
#   3. Search all addresses in a ZIP code

# Usage:
#     python test_location_services.py

# Prerequisites:
#     - AWS credentials configured (aws configure)
#     - An AWS Location Services Place Index created in your account
#     - Set PLACE_INDEX_NAME below or pass via --index flag

# Create a Place Index (one-time):
#     aws location create-place-index \
#         --index-name sov-place-index \
#         --data-source Esri \
#         --region us-east-1
# """

# from __future__ import annotations

# import argparse
# import json
# import sys

# import boto3

# # Default Place Index name — create one in AWS console or via CLI
# DEFAULT_PLACE_INDEX = "sov-place-index"
# DEFAULT_REGION = "us-east-1"


# def get_client(region: str = DEFAULT_REGION, profile: str | None = None):
#     session = boto3.Session(profile_name=profile) if profile else boto3.Session()
#     return session.client("location", region_name=region)


# # -------------------------------------------------------------------------
# # 1. Geocode: Address → Lat/Long
# # -------------------------------------------------------------------------

# def geocode_address(
#     client,
#     address: str,
#     index_name: str = DEFAULT_PLACE_INDEX,
#     max_results: int = 5,
# ) -> list[dict]:
#     """Convert a text address to lat/long coordinates."""
#     resp = client.search_place_index_for_text(
#         IndexName=index_name,
#         Text=address,
#         MaxResults=max_results,
#     )

#     results = []
#     for r in resp.get("Results", []):
#         place = r["Place"]
#         point = place["Geometry"]["Point"]  # [longitude, latitude]
#         results.append({
#             "address": place.get("Label", ""),
#             "longitude": point[0],
#             "latitude": point[1],
#             "country": place.get("Country", ""),
#             "region": place.get("Region", ""),
#             "sub_region": place.get("SubRegion", ""),
#             "municipality": place.get("Municipality", ""),
#             "postal_code": place.get("PostalCode", ""),
#             "street": place.get("Street", ""),
#             "address_number": place.get("AddressNumber", ""),
#             "relevance": r.get("Relevance", 0),
#         })

#     return results


# # -------------------------------------------------------------------------
# # 2. Reverse Geocode: Lat/Long → Address
# # -------------------------------------------------------------------------

# def reverse_geocode(
#     client,
#     latitude: float,
#     longitude: float,
#     index_name: str = DEFAULT_PLACE_INDEX,
#     max_results: int = 5,
# ) -> list[dict]:
#     """Convert lat/long coordinates back to address."""
#     resp = client.search_place_index_for_position(
#         IndexName=index_name,
#         Position=[longitude, latitude],  # Note: [lon, lat] order
#         MaxResults=max_results,
#     )

#     results = []
#     for r in resp.get("Results", []):
#         place = r["Place"]
#         point = place["Geometry"]["Point"]
#         results.append({
#             "address": place.get("Label", ""),
#             "longitude": point[0],
#             "latitude": point[1],
#             "country": place.get("Country", ""),
#             "region": place.get("Region", ""),
#             "municipality": place.get("Municipality", ""),
#             "postal_code": place.get("PostalCode", ""),
#             "street": place.get("Street", ""),
#             "distance": r.get("Distance", 0),
#         })

#     return results


# # -------------------------------------------------------------------------
# # 3. Search by ZIP Code: Find addresses in a postal code area
# # -------------------------------------------------------------------------

# def search_by_zipcode(
#     client,
#     zip_code: str,
#     country: str = "USA",
#     index_name: str = DEFAULT_PLACE_INDEX,
#     max_results: int = 10,
# ) -> list[dict]:
#     """Find addresses/places within a ZIP code area.

#     Uses the ZIP code as a text search with country filter to find
#     locations in that postal code area.
#     """
#     # First, geocode the ZIP to get its center point
#     zip_results = geocode_address(client, f"{zip_code}, {country}", index_name, max_results=1)
#     if not zip_results:
#         print(f"  Could not find ZIP code: {zip_code}")
#         return []

#     center = zip_results[0]
#     print(f"  ZIP {zip_code} center: lat={center['latitude']}, lon={center['longitude']}")

#     # Now search around that point for nearby places
#     resp = client.search_place_index_for_position(
#         IndexName=index_name,
#         Position=[center["longitude"], center["latitude"]],
#         MaxResults=max_results,
#     )

#     results = []
#     for r in resp.get("Results", []):
#         place = r["Place"]
#         point = place["Geometry"]["Point"]
#         pc = place.get("PostalCode", "")
#         results.append({
#             "address": place.get("Label", ""),
#             "longitude": point[0],
#             "latitude": point[1],
#             "postal_code": pc,
#             "municipality": place.get("Municipality", ""),
#             "region": place.get("Region", ""),
#             "distance_meters": r.get("Distance", 0),
#         })

#     return results


# # -------------------------------------------------------------------------
# # CLI
# # -------------------------------------------------------------------------

# def parse_args():
#     p = argparse.ArgumentParser(description="Test AWS Location Services")
#     p.add_argument("--index", default=DEFAULT_PLACE_INDEX, help="Place Index name")
#     p.add_argument("--region", default=DEFAULT_REGION, help="AWS region")
#     p.add_argument("--profile", default=None, help="AWS CLI profile")
#     return p.parse_args()


# def main():
#     args = parse_args()
#     client = get_client(region=args.region, profile=args.profile)
#     index = args.index

#     print("=" * 70)
#     print("  AWS Location Services Test")
#     print("=" * 70)

#     # # --- Test 1: Geocode addresses ---
#     # test_addresses = [
#     #     "500 Industrial Blvd, Chicago, IL 60601, USA",
#     #     "Ambrosetti 2431, Munro, Buenos Aires, Argentina",
#     #     "45 High Street, London SW1A 1AA, United Kingdom",
#     #     "Gewerbering 2, 2440 Moosbrunn, Austria",
#     #     "Rua Barra Longa 11, Betim, Minas Gerais, Brazil",
#     # ]

#     # print("\n--- 1. GEOCODE: Address → Lat/Long ---\n")
#     # for addr in test_addresses:
#     #     print(f"  Input: {addr}")
#     #     try:
#     #         results = geocode_address(client, addr, index)
#     #         if results:
#     #             r = results[0]
#     #             print(f"  Result: lat={r['latitude']:.6f}, lon={r['longitude']:.6f}")
#     #             print(f"  Resolved: {r['address']}")
#     #             print(f"  Country: {r['country']}, Region: {r['region']}, "
#     #                   f"City: {r['municipality']}, ZIP: {r['postal_code']}")
#     #         else:
#     #             print("  No results found")
#     #     except Exception as e:
#     #         print(f"  ERROR: {e}")
#     #     print()

#     # # --- Test 2: Reverse geocode ---
#     # test_coords = [
#     #     (40.7128, -74.0060, "New York City"),
#     #     (41.8781, -87.6298, "Chicago"),
#     #     (51.5074, -0.1278, "London"),
#     # ]

#     # print("\n--- 2. REVERSE GEOCODE: Lat/Long → Address ---\n")
#     # for lat, lon, label in test_coords:
#     #     print(f"  Input: lat={lat}, lon={lon} ({label})")
#     #     try:
#     #         results = reverse_geocode(client, lat, lon, index)
#     #         if results:
#     #             for i, r in enumerate(results[:3]):
#     #                 print(f"  [{i+1}] {r['address']} (dist: {r['distance']:.0f}m)")
#     #         else:
#     #             print("  No results found")
#     #     except Exception as e:
#     #         print(f"  ERROR: {e}")
#     #     print()

#     # --- Test 3: Search by ZIP ---
#     test_zips = [
#         ("147001", "INDIA"),
#         ("10001", "USA"),
#         ("SW1A 1AA", "GBR"),
#     ]

#     print("\n--- 3. SEARCH BY ZIP CODE ---\n")
#     for zip_code, country in test_zips:
#         print(f"  ZIP: {zip_code} ({country})")
#         try:
#             results = search_by_zipcode(client, zip_code, country, index, max_results=5)
#             if results:
#                 for i, r in enumerate(results):
#                     print(f"  [{i+1}] {r['address']} (ZIP: {r['postal_code']}, "
#                           f"dist: {r['distance_meters']:.0f}m)")
#             else:
#                 print("  No results found")
#         except Exception as e:
#             print(f"  ERROR: {e}")
#         print()

#     print("=" * 70)
#     print("  Done")
#     print("=" * 70)


# if __name__ == "__main__":
#     main()

import boto3
import json

# ── Configuration ──────────────────────────────────────────────────────────────
AWS_REGION      = "us-east-1"          # Change to your AWS region
PLACE_INDEX_NAME = "sov-place-index"
# 
DEFAULT_PLACE_INDEX = "sov-place-index"
# "      # Change to your Place Index name
ZIP_CODE        = "32663-330"          # Brazilian ZIP code to look up
# ───────────────────────────────────────────────────────────────────────────────


def search_by_zipcode(client, zip_code: str) -> dict | None:
    """Search for an address using a Brazilian ZIP code (CEP)."""
    try:
        response = client.search_place_index_for_text(
            IndexName=PLACE_INDEX_NAME,
            Text=zip_code,
            FilterCountries=["BRA"],   # Restrict results to Brazil
            MaxResults=5,
        )
        results = response.get("Results", [])
        if not results:
            print(f"No results found for ZIP code: {zip_code}")
            return None
        return results

    except client.exceptions.ResourceNotFoundException:
        print(f"ERROR: Place Index '{PLACE_INDEX_NAME}' not found.")
        print("Create one in the AWS Console → Location Service → Place indexes.")
        return None
    except Exception as e:
        print(f"ERROR: {e}")
        return None


def format_result(result: dict) -> None:
    """Pretty-print a single place result."""
    place = result.get("Place", {})

    label        = place.get("Label", "N/A")
    street       = place.get("AddressNumber", "") + " " + place.get("Street", "")
    municipality = place.get("Municipality", "N/A")
    sub_region   = place.get("SubRegion", "N/A")   # neighbourhood / district
    region       = place.get("Region", "N/A")       # state (e.g. Minas Gerais)
    postal_code  = place.get("PostalCode", "N/A")
    country      = place.get("Country", "N/A")
    geometry     = place.get("Geometry", {}).get("Point", [])

    print("-" * 55)
    print(f"  Full label    : {label}")
    print(f"  Street        : {street.strip() or 'N/A'}")
    print(f"  Neighbourhood : {sub_region}")
    print(f"  City          : {municipality}")
    print(f"  State         : {region}")
    print(f"  ZIP / CEP     : {postal_code}")
    print(f"  Country       : {country}")
    if geometry:
        print(f"  Coordinates   : lat={geometry[1]:.6f}, lon={geometry[0]:.6f}")
    print("-" * 55)


def main():
    # Build the boto3 client (uses ~/.aws/credentials or IAM role automatically)
    client = boto3.client("location", region_name=AWS_REGION)

    print(f"\n🔍 Searching AWS Location Service for ZIP code: {ZIP_CODE}\n")
    results = search_by_zipcode(client, ZIP_CODE)

    if results:
        print(f"Found {len(results)} result(s):\n")
        for idx, result in enumerate(results, start=1):
            print(f"Result #{idx}")
            format_result(result)

        # Optionally dump raw JSON for debugging
        print("\n📄 Raw JSON response:")
        print(json.dumps(results, indent=2, default=str))


if __name__ == "__main__":
    main()