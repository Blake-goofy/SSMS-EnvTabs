# SSMS EnvTabs Wiki

This directory contains the comprehensive documentation for SSMS EnvTabs.

## About This Wiki

These files are designed to be published to the GitHub Wiki for this repository. They can be:
1. Manually copied to the GitHub Wiki interface
2. Used directly from this directory as reference documentation
3. Published via the GitHub Wiki Git repository

## Wiki Pages

- **[Home](Home.md)** - Wiki index and navigation
- **[Installation Guide](Installation-Guide.md)** - Complete installation instructions
- **[Configuration Guide](Configuration-Guide.md)** - Detailed configuration reference
- **[Color Reference](Color-Reference.md)** - All 16 available colors with examples
- **[Wildcard Patterns](Wildcard-Patterns.md)** - SQL LIKE pattern matching guide
- **[Troubleshooting](Troubleshooting.md)** - Common issues and solutions

## Publishing to GitHub Wiki

To publish these pages to the GitHub Wiki:

### Method 1: Via GitHub Web Interface (Easiest)

1. Go to the repository Wiki tab on GitHub
2. Create a new page for each `.md` file
3. Copy the content from each file
4. Save each page

### Method 2: Via Wiki Git Repository

1. Clone the wiki repository:
   ```bash
   git clone https://github.com/Blake-goofy/SSMS-EnvTabs.wiki.git
   ```

2. Copy files from this directory to the wiki repo:
   ```bash
   cp wiki/*.md SSMS-EnvTabs.wiki/
   ```

3. Commit and push:
   ```bash
   cd SSMS-EnvTabs.wiki
   git add .
   git commit -m "Add comprehensive wiki documentation"
   git push
   ```

## File Organization

- All files use GitHub Flavored Markdown (GFM)
- Internal wiki links use the format: `[Link Text](Page-Name.md)`
- External links use full URLs
- Each page is self-contained but cross-references related pages

## Updating the Wiki

When making changes:
1. Update the `.md` files in this directory
2. Commit changes to the main repository
3. Re-publish updated files to the GitHub Wiki (using either method above)
