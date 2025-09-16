# Google Calendar CSV Generator & Editor

A Python desktop application for creating, editing, and managing Google Calendar CSV files with recurring events and course schedules.

## Features

- **Create recurring events**: Define weekly sessions with specific times and locations
- **Load existing CSV files**: Import and modify existing Google Calendar exports
- **Course management**: Organize events by subject with instructor information
- **Session patterns**: Support for complex weekly schedules (multiple sessions per week)
- **CSV export**: Generate Google Calendar-compatible CSV files
- **Configuration saving**: Save and load your event configurations
- **Preview functionality**: Review generated events before export
- **Date format support**: Handle both DD/MM/YYYY and MM/DD/YYYY formats

## Use Cases

- **University students**: Manage course schedules with lectures, labs, and tutorials
- **Teachers**: Create semester-long class schedules
- **Organizations**: Plan recurring meetings and events
- **Personal scheduling**: Organize weekly activities and commitments

## Installation

### Requirements

- Python 3.6 or higher
- tkinter (usually included with Python)
- No additional dependencies required

### Setup

1. Clone the repository:
```bash
git clone https://github.com/yourusername/google-calendar-generator.git
cd google-calendar-generator
```

2. Run the application:
```bash
python calendar_generator.py
```

## Usage

### Creating a New Course/Event

1. **Enter course details**:
   - Course name
   - Instructor/description
   - Start and end dates (DD/MM/YYYY format)

2. **Add sessions**:
   - Click "Aggiungi sessione" (Add session)
   - Select day of the week
   - Set start and end times (HH:MM format)
   - Specify location/room

3. **Configure options**:
   - Private event (default: yes)
   - All-day event (default: no)

4. **Add to events list**:
   - Click "Aggiungi Corso" (Add Course)
   - Course appears in the events table

### Importing Existing CSV Files

1. Click "Carica CSV esistente" (Load existing CSV)
2. Select your Google Calendar CSV export
3. Choose to replace or merge with existing events
4. Events are automatically parsed and loaded

### Editing Events

1. Select an event from the table
2. Click "Modifica" (Edit)
3. Event data loads into the form
4. Make changes and click "Aggiungi Corso" (Add Course)

### Generating CSV Files

1. Click "GENERA CSV" (Generate CSV)
2. Choose save location and filename
3. Import the file into Google Calendar:
   - Go to calendar.google.com
   - Settings → Import & Export
   - Select and import your CSV file

## File Formats

### Input CSV Format
The application reads Google Calendar CSV exports with these columns:
- Subject, Start Date, Start Time, End Date, End Time
- All Day Event, Description, Location, Private

### Output CSV Format
Generates Google Calendar-compatible CSV files with proper formatting for recurring events.

## Configuration Files

Save your event configurations as JSON files for reuse:
- File → Save Configuration
- File → Load Configuration

## Technical Details

### Session Pattern Recognition
The application intelligently groups recurring events from CSV imports, recognizing that multiple calendar entries represent the same weekly session pattern.

### Date Handling
- Supports both Italian (DD/MM/YYYY) and American (MM/DD/YYYY) date formats
- Automatically converts between formats for display and export
- Handles Google Calendar's MM/DD/YYYY CSV format

### Event Calculation
Calculates total event occurrences based on:
- Course duration (start to end date)
- Number of weekly sessions
- Excludes potential holiday periods

## Screenshots

[Add screenshots of the main interface, session dialog, and CSV preview]

## Known Limitations

- Does not handle holiday exclusions automatically
- Limited to weekly recurring patterns
- Requires manual timezone handling

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Future Enhancements

- Holiday calendar integration
- Custom recurrence patterns (bi-weekly, monthly)
- Multiple timezone support
- Batch CSV processing
- Web interface version

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter issues or have questions:
1. Check the [Issues](https://github.com/yourusername/google-calendar-generator/issues) page
2. Create a new issue with:
   - Description of the problem
   - Steps to reproduce
   - Your operating system and Python version
   - Sample CSV file (if relevant)

## Acknowledgments

- Built with Python's tkinter for cross-platform compatibility
- Designed for Google Calendar CSV format compatibility
- Optimized for academic scheduling use cases
