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
    excel-shred -f csv input_dir_a

    \b
    excel-shred -f json -o output input_dir_a
    """

    click.clear()
    click.secho(f"Excel shredding all files to {format} to {outdir}", fg='blue')

    # ensure output directory exists
    if not os.path.exists(outdir):
        os.makedirs(outdir)

    for path in input_dirs:
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
                new_file = exutil.prep_file_name(PurePath(path).name, PurePath(file).name)
                new_path = os.path.join(outdir, new_file)
                shutil.move(file, new_path)


cli()
