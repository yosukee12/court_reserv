# -*- coding: utf-8 -*-

import sys
from pathlib import Path


PACKAGE_NAME = "court_reserv"
PACKAGE_VERSION = "0.1.0"
PACKAGE_DESCRIPTION = "Court Reservation for FFTC"
PACKAGE_AUTHOR = "Nakagawa Yosuke"
PACKAGE_AUTHOR_EMAIL = "ysk3821@gmail.com"


def _read_text(path):
    return Path(path).read_text(encoding="utf-8")


def _find_packages(exclude=()):
    exclude = set(exclude or ())
    packages = []
    for init_path in Path(".").rglob("__init__.py"):
        package = str(init_path.parent).replace("/", ".")
        if package == ".":
            continue
        if any(package == item or package.startswith(f"{item}.") for item in exclude):
            continue
        packages.append(package)
    return sorted(set(packages))


def _run_setup():
    try:
        from setuptools import setup
    except ModuleNotFoundError:
        raise ModuleNotFoundError(
            "setuptools is required for packaging commands other than --name"
        )

    setup(
        name=PACKAGE_NAME,
        version=PACKAGE_VERSION,
        description=PACKAGE_DESCRIPTION,
        long_description=_read_text("README.md"),
        long_description_content_type="text/markdown",
        author=PACKAGE_AUTHOR,
        author_email=PACKAGE_AUTHOR_EMAIL,
        url="",
        license=_read_text("LICENSE"),
        packages=_find_packages(exclude=("tests", "docs")),
    )


if __name__ == "__main__":
    if "--name" in sys.argv:
        print(PACKAGE_NAME)
        raise SystemExit(0)
    _run_setup()
