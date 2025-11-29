# i2i Teaming Tool

Google Apps Script library for managing project workflows with automated folder creation, calendar events, and notifications.

## Architecture

**Thin Client + Central Library**
- `i2iTeamingTool/` - Central library containing all business logic
- `client/` - Thin client script that delegates to the library
- `setup_wizard/` - Setup automation script

## Structure

- **`i2iTeamingTool/`** - Main library
  - `core/` - Execution context, utilities, validation
  - `data/` - Data models (Project, ProjectSheet, Config, Directory)
  - `services/` - Business logic (ProjectService, NotificationService, FormService, etc.)
  - `Main.js` - Public API functions

- **`client/`** - Client script template for districts
- **`planning/`** - Architecture docs and specifications
