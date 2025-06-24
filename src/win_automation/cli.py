"""Command line interface for Windows automation toolkit."""

from typing import Annotated

import typer

app = typer.Typer(
    name="win-automation",
    help="Windows automation toolkit for office and system tasks",
    no_args_is_help=True,
)


@app.command()
def run(
    script_name: Annotated[str, typer.Argument(help="Name of the script to run")],
    args: Annotated[
        list[str] | None, typer.Argument(help="Arguments to pass to the script")
    ] = None,
) -> None:
    """Run a user script with optional arguments."""
    if args is None:
        args = []

    typer.echo(f"Running script: {script_name}")
    typer.echo(f"Arguments: {args}")

    # TODO: Implement script discovery and execution
    typer.echo("Script execution not implemented yet")


@app.command()
def list_scripts() -> None:
    """List all available user scripts."""
    typer.echo("Available scripts:")
    # TODO: Implement script discovery
    typer.echo("No scripts found (discovery not implemented yet)")


def main() -> None:
    """Main entry point for the CLI."""
    app()


if __name__ == "__main__":
    main()
