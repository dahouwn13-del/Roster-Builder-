# CIEL Smart Roster Generator

## Start on Windows

1. Install Python 3 if it is not already installed.
2. Double-click `Start_CIEL_Roster.bat`.
3. The app opens automatically at `http://127.0.0.1:8765`.

## Workflow

1. Select the week starting date.
2. Enter occupancy, arrivals, departures, group movements, and remarks for each day.
3. Review team balances under **Team & balances**.
4. Select **Generate roster**.
5. Modify any shift directly in the roster table.
6. Select **Export Excel** to download the roster in the supplied format.

The generator excludes the management rows requested, keeps Vivek on night duty, uses Shubham as Vivek's night relief, alternates one/two weekly OFF patterns, avoids night-to-morning transitions, and uses 9-hour 30-minute shifts.
