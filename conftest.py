"""Pytest config: register the --integration flag.

By default, the full test suite runs offline. Tests marked
@pytest.mark.integration hit the live network and are skipped unless
``pytest --integration`` is passed.

Rationale: the offline tests verify our parsing logic, but they can't
detect when an upstream site (Wikipedia, Yahoo, dataroma) changes its
page structure. Integration tests catch that class of failure — the kind
that produced the "Could not locate constituent table for NIKKEI225"
button error in May 2026.
"""
import pytest


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "integration: tests that hit the live network (Wikipedia / Yahoo / dataroma)",
    )


def pytest_addoption(parser):
    parser.addoption(
        "--integration",
        action="store_true",
        default=False,
        help="Run live-network integration tests against Wikipedia / Yahoo / dataroma.",
    )


def pytest_collection_modifyitems(config, items):
    if config.getoption("--integration"):
        return
    skip_marker = pytest.mark.skip(
        reason="needs --integration (hits live network — Wikipedia / Yahoo / dataroma)"
    )
    for item in items:
        if "integration" in item.keywords:
            item.add_marker(skip_marker)
