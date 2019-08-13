import os
import click
import shutil
import exutil
from pathlib import PurePath

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'], ignore_unknown_options=True)


@click.command(context_settings=CONTEXT_SETTINGS)
@click.option('--format', '-f', default='all', type=click.Choice(['json', 'csv', 'all']))
@click.option('--outdir', '-o', prompt=True, type=click.Path(), default=lambda: os.getcwd() + '/out')
@click.argument('input_dirs', nargs=-1, type=click.Path(exists=True))
def cli(format, outdir, input_dirs):
    """
    excel-shred  Version 1.0.0

    Open an Excel workbook, and convert all sheets to json datasets
    :param outdir: output directory for files
    :param format: the output format
    :param input_dirs: one or more directory paths containing excel workbooks

    Examples:

    \b
    excel-shred input_dir_a [input_dir_b]

    \b
    excel-shred -f csv input_dir_a [input_dir_b]

    \b
    excel-shred -f csv -o .\output input_dir_a [input_dir_b]
    """

    click.clear()
    click.secho(f"Excel shredding all files to {format} to {outdir}", fg='blue')

    # ensure output directory exists
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    for path in input_dirs:
        max_path = os.path.join(outdir, PurePath(path).name)
        if not os.path.exists(max_path):
            os.makedirs(max_path)

        # find and rip all excel files in all input directories
        files = list(exutil.open_dir(path, ['xls', 'xlsx']))
        count = len(files)
        with(click.progressbar(files, label=f'Shredding {count} files from {path}', length=count)) as bar:
            for file in bar:
                exutil.shred_sheets(file, format)

        # copy all shredded files to output director
        out_files = list(exutil.open_dir(path, ['csv', 'json']))
        out_count = len(out_files)
        with(click.progressbar(out_files, label=f'copying {out_count} output files', length=out_count)) as bar:
            for file in bar:
                new_path = os.path.join(max_path, PurePath(file).name)
                shutil.move(file, new_path)


cli()
