"""Manual integration script for AWS Location Service (place index search by ZIP).

Requires AWS credentials, a Place Index in the same region, and optional
``boto3``. Not run as part of the default unit test suite.

Usage::

    python tests/test_location_services.py
"""

from __future__ import annotations

import json
import logging
import sys
from typing import Any

import boto3

# Configuration — adjust for your account
AWS_REGION = "us-east-1"
PLACE_INDEX_NAME = "sov-place-index"
ZIP_CODE = "32663-330"

logger = logging.getLogger(__name__)


def search_by_zipcode(client: Any, zip_code: str) -> list[dict[str, Any]] | None:
    """Search the place index for results matching ``zip_code`` (e.g. Brazilian CEP)."""
    try:
        response = client.search_place_index_for_text(
            IndexName=PLACE_INDEX_NAME,
            Text=zip_code,
            FilterCountries=["BRA"],
            MaxResults=5,
        )
        results = response.get("Results", [])
        if not results:
            logger.warning("No results found for ZIP code: %s", zip_code)
            return None
        return results

    except client.exceptions.ResourceNotFoundException:
        logger.error("Place Index %r not found.", PLACE_INDEX_NAME)
        logger.error(
            "Create one in the AWS Console: Location Service -> Place indexes.",
        )
        return None
    except Exception as exc:  # pylint: disable=broad-exception-caught
        # Location API may raise various client errors.
        logger.exception("Location search failed: %s", exc)
        return None


def format_result(result: dict[str, Any]) -> None:
    """Log one ``Results[]`` entry from ``search_place_index_for_text``."""
    place = result.get("Place", {})

    label = place.get("Label", "N/A")
    street = place.get("AddressNumber", "") + " " + place.get("Street", "")
    municipality = place.get("Municipality", "N/A")
    sub_region = place.get("SubRegion", "N/A")
    region = place.get("Region", "N/A")
    postal_code = place.get("PostalCode", "N/A")
    country = place.get("Country", "N/A")
    geometry = place.get("Geometry", {}).get("Point", [])

    logger.info("-" * 55)
    logger.info("  Full label    : %s", label)
    logger.info("  Street        : %s", street.strip() or "N/A")
    logger.info("  Neighbourhood : %s", sub_region)
    logger.info("  City          : %s", municipality)
    logger.info("  State         : %s", region)
    logger.info("  ZIP / CEP     : %s", postal_code)
    logger.info("  Country       : %s", country)
    if geometry:
        logger.info(
            "  Coordinates   : lat=%.6f, lon=%.6f",
            geometry[1],
            geometry[0],
        )
    logger.info("-" * 55)


def main() -> None:
    """Run a sample ZIP search and log structured results plus raw JSON."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        stream=sys.stderr,
    )

    client = boto3.client("location", region_name=AWS_REGION)

    logger.info("Searching AWS Location Service for ZIP code: %s", ZIP_CODE)
    results = search_by_zipcode(client, ZIP_CODE)

    if results:
        logger.info("Found %d result(s)", len(results))
        for idx, result in enumerate(results, start=1):
            logger.info("Result #%d", idx)
            format_result(result)

        logger.info("Raw JSON response:\n%s", json.dumps(results, indent=2, default=str))


if __name__ == "__main__":
    main()
