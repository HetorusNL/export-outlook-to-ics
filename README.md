# Export Outlook to ICS

Outlook ICS exporter, for sharing free/busy information with your other calendars and optionally scp the calendar file to a server / another machine.

## Installation

Run the following:

```bash
poetry install --no-root
```

## Usage

To also scp the calendar file (using WSL) to path: `SCP_DST` on host: `SCP_HOST`, make sure to provide the following two environment variables:

| Name     | Description                       |
| -------- | --------------------------------- |
| SCP_HOST | Hostname or IP address to copy to |
| SCP_DST  | /path/to/calendar.ics             |

Run the following:

```bash
poetry run python main.py
```

To run this repeatedly every x seconds, run:

```bash
poetry run python main.py loop  # runs every 900 seconds
poetry run python main.py loop 123  # runs every 123 seconds, substitute 123 for any integer value
```

## License

MIT License, Copyright (c) 2024 Tim Klein Nijenhuis <tim@hetorus.nl>

## Acknowledgements

MIT License, Copyright (c) 2023 Tom Smeets
