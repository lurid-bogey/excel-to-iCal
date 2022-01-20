# excel-to-iCal

The script will read an Excel file and will generate iCal events for every date. The resulting iCal files can be imported in your favourite calendar app.

The first row must contain the name of the event (the summary). The second row can contain an optional description.

![Excel file](docs/input-file.jpg)

```
BEGIN:VEVENT
UID:2022-Rosenmontag@lurid_bogey_ical_generator.dummy.local
TRANSP:TRANSPARENT
X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC
DTSTART;VALUE=DATE:20220228
DTEND;VALUE=DATE:20220301
SUMMARY:Rosenmontag
END:VEVENT
...
```
