# SharePoint Framework Demo Solution

A comprehensive SharePoint Framework (SPFx) solution built with React that demonstrates various SharePoint integration capabilities including CRUD operations, list management, and user profile information.

## Features

- **Lists Management**: View all SharePoint lists in the current site
- **CRUD Operations**: Create, read, update, and delete list items
- **User Profile**: Display current user and site information
- **Modern UI**: Built with Fluent UI React components for a consistent Microsoft 365 experience
- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **TypeScript**: Fully typed for better development experience

## Prerequisites

- [Node.js](https://nodejs.org/) (v18.x or v16.x)
- [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- [Yeoman](https://yeoman.io/) and [@microsoft/generator-sharepoint](https://www.npmjs.com/package/@microsoft/generator-sharepoint)

## Getting Started

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure SharePoint Site

The solution is currently configured for the SharePoint site: **https://pjqd.sharepoint.com**

To change this, update the `config/serve.json` file:

```json
{
	"$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
	"port": 4321,
	"https": true,
	"initialPage": "https://yourdomain.sharepoint.com/_layouts/workbench.aspx"
}
```

### 3. Build and Test Locally

```bash
# Build the solution
npm run build

# Start the local development server
gulp serve
```

This will open the SharePoint Workbench in your browser where you can test the web part.

### 4. Package for Deployment

```bash
# Build for production
gulp build --ship

# Package the solution
gulp package-solution --ship
```

The packaged solution will be available in the `sharepoint/solution` folder as a `.sppkg` file.

### 5. Deploy to SharePoint

1. Go to your SharePoint Admin Center
2. Navigate to "More features" > "Apps" > "App Catalog"
3. Upload the `.sppkg` file to the App Catalog
4. Click "Deploy" to make it available across your tenant
5. Go to any SharePoint site where you want to use the web part
6. Add the web part to a page through the web part picker

## Project Structure

```
src/
├── webparts/
│   └── demo/
│       ├── components/
│       │   ├── Demo.tsx           # Main component with navigation
│       │   ├── ListViewer.tsx     # Component for viewing SharePoint lists
│       │   ├── ItemViewer.tsx     # Component for viewing and managing list items
│       │   └── UserInfo.tsx       # Component for displaying user and site info
│       ├── models/
│       │   └── IListItem.ts       # TypeScript interfaces for data models
│       ├── services/
│       │   └── SharePointService.ts # Service for SharePoint REST API operations
│       └── DemoWebPart.ts         # Main web part class
```

## Key Components

### SharePointService

Handles all SharePoint REST API operations including:

- Fetching lists and list items
- Creating new list items
- Getting user and site information
- Error handling and type safety

### Demo Component

Main component featuring:

- Tabbed navigation using Fluent UI Pivot
- State management for different views
- Integration with SharePoint services

### ListViewer Component

- Displays all SharePoint lists in a table format
- Search functionality
- Click to drill down into list items

### ItemViewer Component

- Shows items from a selected list
- Create new items with a slide-out panel
- Back navigation to lists view

### UserInfo Component

- Displays current user profile information
- Shows site metadata and information

## Permissions Required

The web part requires the following SharePoint permissions:

- **Read**: To view lists and list items
- **Write**: To create new list items
- **User Profile**: To read current user information

## Browser Support

- Microsoft Edge (Chromium-based)
- Google Chrome
- Mozilla Firefox
- Safari (latest versions)

## Development Commands

```bash
# Install dependencies
npm install

# Start development server
gulp serve

# Build for development
gulp build

# Build for production
gulp build --ship

# Package solution
gulp package-solution

# Clean build artifacts
gulp clean

# Run tests (if available)
npm test
```

## Troubleshooting

### Common Issues

1. **CORS Errors**: Ensure you're accessing the SharePoint site with proper authentication
2. **Permission Denied**: Check that your user account has appropriate permissions on the SharePoint site
3. **Build Errors**: Ensure all dependencies are installed with `npm install`

### Debugging

1. Use browser developer tools to inspect network requests
2. Check the console for any JavaScript errors
3. Verify the SharePoint REST API endpoints are accessible

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License.

## Learn More

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- [Fluent UI React Components](https://developer.microsoft.com/en-us/fluentui)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
