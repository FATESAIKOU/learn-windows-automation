"""Command line interface for Windows automation toolkit."""

from typing import Annotated

import typer

from .core import ConfigManager, ScriptManager
from .core.exceptions import AutomationError, ScriptNotFoundError

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

    script_manager = ScriptManager()

    try:
        typer.echo(f"Running script: {script_name}")
        if args:
            typer.echo(f"Arguments: {args}")

        exit_code = script_manager.execute_script(script_name, args)

        if exit_code == 0:
            typer.echo("Script completed successfully", color=typer.colors.GREEN)
        else:
            typer.echo(
                f"Script failed with exit code: {exit_code}",
                color=typer.colors.RED,
            )
            raise typer.Exit(exit_code)

    except ScriptNotFoundError as e:
        typer.echo(f"Error: {e}", color=typer.colors.RED)
        raise typer.Exit(1) from None
    except AutomationError as e:
        typer.echo(f"Execution error: {e}", color=typer.colors.RED)
        raise typer.Exit(1) from None


@app.command()
def list_scripts() -> None:
    """List all available user scripts."""
    script_manager = ScriptManager()
    config_manager = ConfigManager()

    scripts = script_manager.discover_scripts()

    if not scripts:
        typer.echo("No scripts found in configuration")
        return

    typer.echo("Available scripts:")
    for script_name in scripts:
        script_info = config_manager.get_script_info(script_name)
        if script_info:
            description = script_info.get("description", "No description")
            category = script_info.get("category", "uncategorized")
            enabled = script_info.get("enabled", True)

            status = "✓" if enabled else "✗"
            typer.echo(f"  {status} {script_name} ({category}) - {description}")


@app.command()
def info(
    script_name: Annotated[str, typer.Argument(help="Name of the script")]
) -> None:
    """Show detailed information about a script."""
    config_manager = ConfigManager()
    script_manager = ScriptManager()

    script_info = config_manager.get_script_info(script_name)
    if not script_info:
        typer.echo(f"Script '{script_name}' not found", color=typer.colors.RED)
        raise typer.Exit(1)

    # Validate script exists
    is_valid = script_manager.validate_script(script_name)

    typer.echo(f"Script: {script_name}")
    typer.echo(f"Description: {script_info.get('description', 'No description')}")
    typer.echo(f"Category: {script_info.get('category', 'uncategorized')}")
    typer.echo(f"Subcategory: {script_info.get('subcategory', 'none')}")
    typer.echo(f"Path: {script_info.get('path')}")
    typer.echo(f"Enabled: {script_info.get('enabled', True)}")
    typer.echo(f"Valid: {'Yes' if is_valid else 'No'}")


def main() -> None:
    """Main entry point for the CLI."""
    app()


if __name__ == "__main__":
    main()
