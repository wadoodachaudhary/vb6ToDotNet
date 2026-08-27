# HomeFront Suite Code Documentation

This directory contains the interactive, Confluence-style offline documentation system for the HomeFront codebase (including VB6 HFSystem, HFEst, HFPayables, AutoNotice, and the modern C# Blazor HomeFrontPB).

## How to View in a Browser

### Option 1: Direct File Access (No Setup)
Simply open the `index.html` file in your preferred web browser.
- **On macOS**: Double-click `index.html` or drag it into Safari/Chrome.
- **Via Terminal**:
  ```bash
  open index.html
  ```

### Option 2: Run a Local Dev Server
If you prefer to serve the files via HTTP:
- **Using Python 3**:
  ```bash
  python3 -m http.server 8000
  ```
  Then open [http://localhost:8000](http://localhost:8000) in your browser.

- **Using Node.js**:
  ```bash
  npx http-server -p 8000
  ```
  Then open [http://localhost:8000](http://localhost:8000) in your browser.

## Features Built-in
- **Interactive Sidebar**: Full tree-view hierarchy of all HomeFront systems.
- **Fast Live Search**: Search across topics, files, database tables, and migration statuses in real-time.
- **Dark Mode / Light Mode**: Instant toggle for comfortable reading.
- **Rich Callouts & Tables**: Clearly structured schemas, migration registries, and architectural walkthroughs.
- **Responsive Layout**: Designed for seamless desktop and mobile viewports.
