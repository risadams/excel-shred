import click
import exutil

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'], ignore_unknown_options=True)


@click.command(context_settings=CONTEXT_SETTINGS)
@click.argument('files', nargs=-1, type=click.Path(exists=True))
def cli(files):
    """
    Open an Excel workbook, and convert all sheets to json datasets
    :param files: an excel workbook filename

    Example:
    \b
    excel-shred filename
    """

    for file in files:
        print(f'excel-shred:  operating file: {file}')
        exutil.openSheets(file)


cli()
