import subprocess
import tkinter as tk
from collections import deque
from pathlib import Path
from tkinter import filedialog
from typing import List, Union

import click


PATH_TO_CONVERTER = Path(r"C:\Program Files\Garmin\VIRB Edit\GMetrixConverter.exe")
SORT_ORDER = "eByType"  # -s flag
FIELDS_TO_EXPORT = {  # -d flag
    "eGenericBegin_4": "Position",
    "eAltitude_5": "Altitude",
    "eTemperature_3": "Temperature",
    "eSpeed_4": "2D Speed (Recorded)",
    "eSpeed_5": "2D Speed (Calculated)",
    "eAcceleration_5": "Acceleration (Filtered)",
    "eGyroscope_4": "Gyroscope",
    "eDistance_5": "Distance",
    "eGrade_5": "Grade",
    "eCourse_4": "Course",
    "eBarometricPressure_5": "Barometric Pressure",
    "eVelocity_5": "3D Velocity (Recorded)",
    "e3dSpeed_4": "3D Speed (Recorded)",
    "e3dSpeed_5": "3D Speed (Calculated)",
    "eRawAltitude_4": "Raw Altitude",
    "eRawBarometricPressure_4": "Raw Barometric Pressure",
    "eAltitudeUncertainty_4": "Altitude Uncertainty",
    "ePositionUncertainty_4": "Position Uncertainty",
    "eRotation": "Rotation",
    "eGPSHighSpeed": "GPS High Speed Tag",
    "eHighAltitude": "High Altitude Tag",
    "eTurn": "Turn Tag",
    "eAltitudeDescent": "Descent Tag",
    "eZeroGravity": "Zero Gravity Tag",
}


def processing_pipeline(datadir: Path) -> None:
    """
    Recursively search for *.fit files contained in the provided directory & convert to CSV.

    Files are first converted to JSON using Garmin's GMetrix converter, then converted to a CSV

    Note: *.fit files with an exactly named *.json partner in the same directory are ignored.
    """
    # Convert unconverted files to JSON & add to CSV conversion queue
    csv_conversion_queue = deque()
    for fit_file in datadir.rglob("*.fit"):
        # Check for existing conversion
        file_as_json = fit_file.with_suffix(".json")
        if file_as_json.exists():
            continue
        else:
            call_converter(fit_file)
            csv_conversion_queue.append(file_as_json)


def build_cli_cmd(
    in_filepath: Path,
    out_filepath: Union[Path, None] = None,
    sort_order: str = SORT_ORDER,
    fields_to_export: List = FIELDS_TO_EXPORT.keys(),
    converter_path: Path = PATH_TO_CONVERTER,
) -> str:
    """Build the CLI command string from the provided parameters to pass to GMetrixConverter."""
    if not out_filepath:
        # If a new output filename is not provided, mirror the input filename
        out_filepath = in_filepath.with_suffix(".json")

    return (
        f'"{converter_path}" -i "{in_filepath}" -o "{out_filepath}" '
        f"-s {sort_order} -d {' '.join(fields_to_export)}"
    )


def call_converter(in_filepath: Path) -> None:
    """Call the GMetrix Converter for the provided filepath."""
    subprocess.run(build_cli_cmd(in_filepath))


def fit_json_to_csv(in_filepath: Path) -> None:
    """
    Convert Garmin Virb data JSON to a CSV.

    Data is output by Garmin's GMetrixConverter as a JSON of the following sample form:
        {
            "metadata": {
                "deviceType": "2687 (Unknown Device)",
                "distance": 8890.5468448859046,
                "duration": 457.98799991607666,
                "source": "2019-08-18-10-17-15.fit",
                "startTime": "2019-8-18T17:17:14.999Z",
                "types": [{"Position": {"units": str}}, ...]
            },
            "typedata": [
                {
                    "type": str,
                    "values": [
                        {
                            "time": str
                            "some_value": many_typed
                        }
                    ]
                }
            ]
        }
    """
    raise NotImplementedError


@click.command()
@click.option("-d", "--datadir", default=None, help="Top level data directory")
def cli(datadir: Union[str, None]):
    """
    CLI Userflow.

    If no `datadir` is explicitly specified, the user is prompted to select the
    top-level data directory.
    """
    if not datadir:
        # Generate a Tk file selection dialog to select the top level data dir
        # if none is provided on the CLI
        root = tk.Tk()
        root.withdraw()
        datadir = Path(filedialog.askdirectory(title="Select Data Directory"))
    else:
        datadir = Path(datadir)

    processing_pipeline(datadir)


if __name__ == "__main__":
    cli()
