import subprocess
import pathlib
from datetime import datetime
import typer

app = typer.Typer(add_completion=False)

def run_cmd(cmd: str):
    print(f"[mailmonkey] {cmd}")
    # Use shell=True so we can `cd && python ...` on Windows
    subprocess.check_call(cmd, shell=True)

@app.command(help="Build -> Generate -> Finalize for a single campaign.")
def run(
    campaign_number: int = typer.Option(..., help="Campaign number, e.g., 1"),
    template_id: int = typer.Option(101, help="Letter template ID"),
    campaign_name: str = typer.Option("Campaign", help="Campaign name stored in outputs"),
    target_size: int = typer.Option(5000, help="Max rows to build"),
    prior_exact: int = typer.Option(0, help="Use 0 for 'never mailed' cohort"),
    strict_150: bool = typer.Option(True, help="Favor 150-piece ZIP5 packing"),
    mandatory: list[str] = typer.Option(..., help="One or more mandatory CSVs (repeat flag)"),
    optional: list[str] = typer.Option([], help="Optional CSVs (repeat flag)"),
    sig_image: str = typer.Option("sig_ed.png", help="Signature image path relative to repo root"),
    name: str = typer.Option("Ed & Albert Beluli", help="Contact name for letters"),
    phone: str = typer.Option("916-905-7281", help="Contact phone"),
    email: str = typer.Option("eabeluli@gmail.com", help="Contact email"),
    root: str = typer.Option(".", help="Repo root (where the scripts live)"),
    debug: bool = typer.Option(True, help="Pass --debug to builder"),
    skip_singles: bool = typer.Option(True, help="Generate combined PDF only"),
):
    rootp = pathlib.Path(root).resolve()

    # Example campaign dir: Campaign_1_Aug2025
    mon_year = datetime.now().strftime("%b%Y")
    campaign_dir_name = f"Campaign_{campaign_number}_{mon_year}"
    campaign_dir = rootp / campaign_dir_name

    # 1) Build
    mand = " ".join(f'"{m}"' for m in mandatory)
    opt = " ".join(f'"{o}"' for o in optional) if optional else ""
    strict = "--strict-150" if strict_150 else ""
    dbg = "--debug" if debug else ""
    prior = f"--prior-exact {prior_exact}" if prior_exact is not None else ""

    build_cmd = (
        f'cd "{rootp}" && '
        f'python build_campaign_timegap.py '
        f'--campaign-name "{campaign_name}" '
        f'--campaign-number {campaign_number} '
        f'--target-size {target_size} '
        f'--mandatory {mand} '
        f'{"--optional " + opt + " " if opt else ""}'
        f'{prior} {strict} {dbg}'
    )
    run_cmd(build_cmd)

    # 2) Generate
    skip = "--skip-singles" if skip_singles else ""
    gen_cmd = (
        f'cd "{campaign_dir}" && '
        f'python ..\\generate_letters.py '
        f'--csv "campaign_master.csv" '
        f'--outdir "Singles" '
        f'--combine-out "letters_batch.pdf" '
        f'--map-out "letters_mapping.csv" '
        f'--template-id {template_id} {skip} '
        f'--sig-image "..\\{sig_image}" '
        f'--name "{name}" '
        f'--phone "{phone}" '
        f'--email "{email}"'
    )
    run_cmd(gen_cmd)

    # 3) Finalize
    fin_cmd = (
        f'cd "{rootp}" && '
        f'python finalize_or_rebuild.py '
        f'--campaign-dir "{campaign_dir_name}" '
        f'--write-marker'
    )
    run_cmd(fin_cmd)

if __name__ == "__main__":
    app()
