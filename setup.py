import sys

from setuptools import find_packages
from setuptools import setup

assert sys.version_info >= (3, 6, 0), "tool requires Python 3.6+"


with open("README.md") as readme_file:
    readme = readme_file.read()

install_requirements = ["Click>=5.0", "lxml", "bs4", "xlsxwriter"]

setup_requirements = ["setuptools_scm"]

setup(
    author="simpleQE Organization",
    author_email="qesimple@gmail.com",
    classifiers=[
        "Natural Language :: English",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3 :: Only",
        "Topic :: Software Development :: Quality Assurance",
        "Topic :: Software Development :: Testing",
    ],
    description="Test case generator tool",
    entry_points={"console_scripts": ["tcgen=src:main"]},
    install_requires=install_requirements,
    long_description=readme,
    long_description_content_type="text/markdown",
    include_package_data=True,
    setup_requires=setup_requirements,
    python_requires=">=3.6",
    use_scm_version=True,
    keywords="test",
    name="tcgen",
    packages=find_packages(include=["src"]),
    url="https://github.com/simpleQE/tcGen.git",
    license="GPLv3",
    zip_safe=False,
)
