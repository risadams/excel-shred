import click
import exutil

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'], ignore_unknown_options=True)


@click.command(context_settings=CONTEXT_SETTINGS)
@click.option('--format','-f', default='json', type=click.Choice(['json', 'csv']))
@click.argument('input_dirs', nargs=-1, type=click.Path(exists=True))
def cli(format, input_dirs):
    """
    Open an Excel workbook, and convert all sheets to json datasets
    :param input_dirs: one or more directory paths containing excel workbooks

    Example:
    \b
    excel-shred input_dir_a [input_dir_b]
    """

    print(f"Excel shredding all files to {format}")

    for path in input_dirs:
        for file in exutil.open_dir(path):
            print(f'\tShredding : {file}')
            exutil.shred_sheets(file, format)


cli()
