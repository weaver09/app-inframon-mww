# InfraMonitor

InfraMonitor is a simple Azure-hosted endpoint monitoring application.

## Architecture

- Azure App Service
- Azure SQL Database
- Azure Functions

## Current Features

- Displays configured monitored endpoints
- Reads endpoint configuration from Azure SQL Database
- Provides a health endpoint for database connectivity

## Planned Features

- Add/edit/delete monitored endpoints
- Scheduled monitoring with Azure Functions
- Response time tracking
- Uptime history
- Managed Identity authentication
- Application Insights