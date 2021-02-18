import click

from src.testcase import generate


@click.version_option("0.1")
@click.group()
def main():
    """Test case generator tool"""
    pass


main.add_command(generate)


if __name__ == "__main__":
    main()