# Power Platform Solution Scraper- Usage Guide

## Overview

This tool analyzes all solutions in a Power Platform environment and creates comprehensive documentation with:

1. **Master Analysis Workbook** - Cross-solution analysis with:
   - All solutions summary with component counts and record totals
   - All tables master list with record counts
   - Component reuse analysis (components in multiple solutions)
   - Table ownership map
   - Environment-wide statistics
   - **Unmanaged components (not in any solution)** - Tables, Flows, Apps, Web Resources, Forms, Option Sets

2. **Individual Solution Workbooks** - Detailed breakdowns for each solution with:
   - Solution summary metadata
   - **All Components Summary with logical names** (searchable, no more GUID hunting!)
   - **Component-specific sheets** for each type found (Tables, Flows, Views, Roles, Web Resources, Forms, Option Sets, etc.)
   - Detailed metadata for major component types

**Key Features:**
- ✅ **Logical names and schema names** for all components (find components by name, not GUID)
- ✅ **Actual record counts** using FetchXML aggregate (unlimited, works with millions of records)
- ✅ **Component-specific sheets** with detailed metadata (flows, views, roles, forms, web resources, etc.)
- ✅ **Unmanaged component detection** - Find orphaned components not in any solution
- ✅ **In-memory data collection** for fast cross-solution analysis
- ✅ **Environment-dated output folders** for multi-environment tracking
- ✅ **Professional Excel formatting** with auto-sized columns

All data is collected once and stored in memory for efficient processing and cross-solution comparisons.

---

## Prerequisites

### Required Software
- **Python 3.7+** installed on your system
- Internet connection
- Access to target Power Platform environment

### Required Python Packages
```bash
pip install msal requests openpyxl
```

Or install all at once:
```bash
pip install msal requests openpyxl --break-system-packages
```

### Required Permissions
- **System Administrator** or **System Customizer** role in the target environment
- Ability to read solution components and metadata
- Ability to authenticate via Azure AD device code flow

---

## Quick Start

### Basic Usage
```bash
python3 solution_exporter_with_master.py --url https://yourorg.crm.dynamics.com
```

### What Happens Next
1. **Authentication Prompt** - You'll see a device code authentication message
2. **Browser Step** - Open the provided URL and enter the code
3. **Sign In** - Authenticate with your Microsoft account
4. **Data Collection** - Script fetches solutions, components, and metadata
5. **File Creation** - Exports are saved to timestamped folder

---

## Command-Line Options

### Required Arguments

**`--url`** - Your Dataverse environment URL
```bash
--url https://yourorg.crm.dynamics.com
--url https://yourorg.crm4.dynamics.com
--url https://yourorg-dev.crm.dynamics.com
```

### Optional Arguments

**`--output-dir`** - Custom output directory (default: `solution_master_export`)
```bash
--output-dir "C:/PowerPlatform/Exports"
--output-dir "/home/user/pp-exports"
```

**`--include-microsoft`** - Include Microsoft base solutions (default: excluded)
```bash
--include-microsoft
```
By default, these solutions are excluded:
- Default*
- Dynamics*
- Microsoft*
- msdyn_*
- msdynce_*
- HCMCommon*

**`--client-id`** - Custom Azure AD Application ID (optional)
```bash
--client-id "your-app-id-here"
```

**`--tenant-id`** - Specific Azure AD Tenant (optional, default: common)
```bash
--tenant-id "your-tenant-id-here"
```

---

## Usage Examples

### Example 1: Basic Export
Export all custom solutions from production environment:
```bash
python3 solution_exporter_with_master.py \
  --url https://contoso.crm.dynamics.com
```

**Output:**
```
solution_master_export/
└── Environment Solution Analysis contoso 20260112/
    ├── MASTER_Analysis_20260112.xlsx
    ├── Location Management/
    │   └── Location Management_2024-10-20.xlsx
    ├── Alerts/
    │   └── Alerts_2024-11-15.xlsx
    └── ...
```

### Example 2: Include Microsoft Solutions
Export all solutions including system solutions:
```bash
python3 solution_exporter_with_master.py \
  --url https://contoso.crm.dynamics.com \
  --include-microsoft
```

### Example 3: Custom Output Location
Export to specific directory:
```bash
python3 solution_exporter_with_master.py \
  --url https://contoso-dev.crm.dynamics.com \
  --output-dir "C:/PowerPlatform/DevExports"
```

### Example 4: Multiple Environments
Export from DEV, TEST, and PROD in sequence:
```bash
# DEV Environment
python3 solution_exporter_with_master.py \
  --url https://contoso-dev.crm.dynamics.com

# TEST Environment  
python3 solution_exporter_with_master.py \
  --url https://contoso-test.crm.dynamics.com

# PROD Environment
python3 solution_exporter_with_master.py \
  --url https://contoso.crm.dynamics.com
```

---

## Authentication Process

### Device Code Flow
The script uses Azure AD device code authentication:

1. **Script displays code:**
   ```
   To sign in, use a web browser to open the page 
   https://microsoft.com/devicelogin and enter the code ABC123XYZ
   ```

2. **Open browser** and navigate to: https://microsoft.com/devicelogin

3. **Enter code** displayed by the script

4. **Sign in** with your Microsoft account (must have access to the environment)

5. **Consent** to permissions if prompted

6. **Return to terminal** - Script continues automatically

### Troubleshooting Authentication

**Issue: "Authentication failed"**
- Ensure you have System Administrator or System Customizer role
- Verify the environment URL is correct
- Check if MFA is required and complete it during browser sign-in

**Issue: "Token expired"**
- Re-run the script - tokens are not cached between runs
- Complete authentication within 15 minutes of receiving the code

---

## Output Structure

### Folder Hierarchy
```
{output-dir}/
└── Environment Solution Analysis {env-name} {YYYYMMDD}/
    ├── MASTER_Analysis_{YYYYMMDD}.xlsx          ← Master workbook
    │
    ├── {Solution Name 1}/                       ← Solution folder
    │   └── {Solution Name 1}_{install-date}.xlsx
    │
    ├── {Solution Name 2}/                       ← Solution folder
    │   └── {Solution Name 2}_{install-date}.xlsx
    │
    └── ...
```

### Example Output
```
solution_master_export/
└── Environment Solution Analysis contoso 20260112/
    ├── MASTER_Analysis_20260112.xlsx
    ├── Location Management/
    │   └── Location Management_2024-10-20.xlsx
    ├── Asset Tracking/
    │   └── Asset Tracking_2024-09-05.xlsx
    └── User Management/
        └── User Management_2024-12-01.xlsx
```

---

## Master Workbook Contents

**File:** `MASTER_Analysis_{YYYYMMDD}.xlsx`

### Sheet 1: All Solutions
Summary of every solution in the environment:
- Solution Name, Unique Name, Version
- Publisher, Type (Managed/Unmanaged)
- Installed On date
- Total Components (count)
- Tables, Flows, Views, Roles (counts)
- **Total Records** - Sum of all table records in solution

**Use Cases:**
- Compare solution sizes
- Identify data-heavy solutions
- Track solution versions
- Migration planning by data impact

### Sheet 2: All Tables
Master list of every table across all solutions:
- Table Name, Logical Name, Schema Name
- Record Count (actual data volume)
- Is Custom (Yes/No)
- In # Solutions (how many solutions contain it)
- Solutions (list of solutions containing it)

**Use Cases:**
- Find tables in multiple solutions (potential duplicates)
- Identify high-data tables for migration prioritization
- Locate empty tables for cleanup
- Data consolidation planning

### Sheet 3: Component Reuse
Components that appear in multiple solutions:
- Component Type
- Component ID
- Used In # Solutions
- Solutions (list)

**Sorted by:** Most reused first

**Use Cases:**
- Identify shared components
- Find candidates for foundation solutions
- Reduce duplication
- Understand cross-solution dependencies

### Sheet 4: Table Ownership
Maps tables to their "primary owner" solution:
- Table Name, Logical Name
- Record Count
- Primary Owner Solution (alphabetically first)
- # Solutions (total containing it)
- All Solutions (list)

**Use Cases:**
- Determine data migration responsibility
- Identify table ownership conflicts
- Plan solution consolidation
- Data governance decisions

### Sheet 5: Statistics
Environment-wide metrics:
- Total Solutions (managed vs unmanaged)
- Total Components
- Total Unique Tables
- Custom Tables vs System Tables
- Tables with Data vs Empty Tables
- **Total Records** (across ALL tables)
- Components Reused Across Solutions

**NEW - Unmanaged Components Section:**
- Unmanaged Tables (custom tables not in any solution)
- Unmanaged Flows (workflows/flows not in any solution)
- Unmanaged Canvas Apps (canvas apps not in any solution)
- Unmanaged Web Resources (JS/HTML/CSS files not in any solution)
- Unmanaged Forms (forms not in any solution)
- Unmanaged Option Sets (global option sets not in any solution)
- **Total Unmanaged Components**

**Use Cases:**
- Environment health check
- Capacity planning
- Executive reporting
- Cleanup prioritization
- **ALM compliance** - Identify orphaned components
- **Pre-migration audit** - Find components that won't be exported

### Sheets 6-11: Unmanaged Components (RED Headers)

**Note:** These sheets only appear if unmanaged components are found.

These sheets identify custom components that are **NOT** in any solution - these are "orphaned" or "loose" components that exist directly in the environment.

#### Sheet 6: Unmanaged Tables
Custom tables not in any solution:
- Table Name, Logical Name, Schema Name
- **Record Count** (shows data impact)
- Primary ID, Metadata ID

**Why This Matters:**
- These tables won't be exported with solutions
- Must be manually migrated or added to solutions
- May represent development/testing artifacts
- Could be outdated or abandoned customizations

#### Sheet 7: Unmanaged Flows
Workflows/flows not in any solution:
- Name, Unique Name, Category, State
- Primary Entity, Created On, Workflow ID

**Common Causes:**
- Flows created directly in the maker portal
- Test flows that were never solutioned
- Personal automation not intended for ALM

#### Sheet 8: Unmanaged Canvas Apps
Canvas apps not in any solution:
- Display Name, Name, Created On, Canvas App ID

**Impact:**
- These apps won't be included in solution exports
- Must be added to solutions for proper ALM

#### Sheet 9: Unmanaged Web Resources
JS/HTML/CSS files not in any solution:
- Name, Display Name, Type, Created On, Web Resource ID

**Common Issues:**
- Leftover files from old customizations
- Test scripts that were never cleaned up
- May cause bloat in environment

#### Sheet 10: Unmanaged Forms
Forms not in any solution:
- Name, Entity, Type (Main, Quick Create, etc.), Created On, Form ID

**Typical Scenario:**
- Forms created during customization testing
- Backup forms that were never deleted
- Forms from disabled features

#### Sheet 11: Unmanaged Option Sets
Global option sets not in any solution:
- Name, Display Name, Metadata ID

**Why Review:**
- Global option sets should typically be in solutions
- Unmanaged option sets complicate migrations
- May indicate incomplete solution setup

**Key Benefits of Unmanaged Component Detection:**
- ✅ **ALM Compliance** - Ensure everything is properly solutioned
- ✅ **Migration Readiness** - Know what won't export with solutions
- ✅ **Environment Cleanup** - Identify orphaned artifacts
- ✅ **Governance** - Track unsolutioned customizations
- ✅ **Accurate Analysis** - Uses the same data collected for solution analysis

---

## Individual Solution Workbooks

**Files:** `{SolutionName}/{SolutionName}_{install-date}.xlsx`

Each solution gets its own folder and workbook. The workbook structure adapts based on what components are in the solution.

### Standard Sheets (Always Present)

#### Sheet 1: Summary
Solution metadata:
- Solution Name, Unique Name, Version
- Publisher, Type, Installed On
- Description
- Total Components

#### Sheet 2: All Components Summary
Complete searchable inventory with logical names:
- **Component Type** (friendly name, e.g., "Workflow/Flow")
- **Type Code** (numeric identifier, e.g., 29)
- **Logical Name** (technical identifier for code/scripts)
- **Schema Name** (friendly name or display name)
- **Object ID** (component GUID)
- **Component ID** (solution component GUID)

**Key Benefits:**
- **Find components by name** - No more GUID hunting!
- **Searchable** - Filter by Type Code (e.g., 29 for flows) then search Logical Name
- **Developer-friendly** - Logical names are what you use in PowerShell, code, and APIs
- **Complete audit trail** - Every component in one place

**Example Usage:**
```
Finding a specific flow:
1. Open "All Components Summary" sheet
2. Filter Type Code column to "29"
3. Search Logical Name column for "dailysync"
4. The Object ID (GUID) is right there if you need it
```

### Dynamic Component Sheets (Based on Solution Content)

**One detailed sheet per component type found in the solution**

The script automatically creates specialized sheets for each component type present. Sheet names follow the format: `{TypeCode}-{TypeName}` (e.g., `1-Entity/Table`, `29-Workflow/Flow`)

#### Tables/Entities (Type 1)
**Sheet Name:** `1-Entity/Table`

**Columns:** Type Code | Table Name | Logical Name | Schema Name | Record Count | Is Custom | Primary ID | Primary Name | Object Type Code | Object ID

**Shows:**
- Display names and logical names (for FetchXML, code)
- **Actual record counts** from the environment
- Custom vs system table identification
- Primary key and name attributes
- Object type codes for reference

**Use Cases:**
- Data migration planning (sort by Record Count)
- Identify custom tables (filter Is Custom = Yes)
- Find tables with/without data

#### Workflows/Flows (Type 29)
**Sheet Name:** `29-Workflow/Flow`

**Columns:** Type Code | Flow Name | Unique Name | Description | Category | Primary Entity | Type | State | Created On | Modified On | Object ID

**Shows:**
- Flow unique names (the logical identifier)
- Categories and workflow types
- Active/inactive state
- Associated entities
- Creation and modification dates

**Use Cases:**
- Document automation workflows
- Identify inactive flows
- Find flows for specific entities

#### Views/Saved Queries (Type 26, 10293)
**Sheet Name:** `26-Saved Query/View` or `10293-View`

**Columns:** Type Code | View Name | Description | Entity | Query Type | Is Default | Is Quick Find | Is Private | Created On | Modified On | Object ID

**Shows:**
- View names and descriptions
- Which table/entity the view is for
- Special view types (default, quick find)
- Public vs private views
- Creation and modification dates

**Use Cases:**
- Document entity views
- Identify default and quick find views
- Find private views for cleanup

#### Security Roles (Type 20)
**Sheet Name:** `20-Role`

**Columns:** Type Code | Role Name | Business Unit ID | Is Customizable | Is Managed | Created On | Modified On | Object ID

**Shows:**
- Role names
- Business unit assignments
- Customization status
- Managed vs unmanaged

**Use Cases:**
- Security audit
- Role documentation
- Identify custom roles

#### Web Resources (Type 61)
**Sheet Name:** `61-Web Resource`

**Columns:** Type Code | Name | Display Name | Description | Type | Created On | Modified On | Object ID

**Shows:**
- File names (logical identifiers)
- Resource types (HTML, CSS, JavaScript, PNG, JPG, GIF, etc.)
- Descriptions
- Creation and modification dates

**Use Cases:**
- Document custom web resources
- Identify JavaScript/HTML files
- Find unused resources

#### System Forms (Type 60)
**Sheet Name:** `60-System Form`

**Columns:** Type Code | Form Name | Description | Type | Entity | State | Created On | Modified On | Object ID

**Shows:**
- Form names
- Form types (Main, Quick Create, Quick View, Card, etc.)
- Associated entities
- Active/inactive state
- Creation and modification dates

**Use Cases:**
- Document custom forms
- Identify forms per entity
- Find inactive forms

#### Option Sets (Type 9)
**Sheet Name:** `9-Option Set`

**Columns:** Type Code | Name | Display Name | Description | Is Global | Object ID

**Shows:**
- Option set names (logical identifiers)
- Display names
- Descriptions
- Global vs local scope

**Use Cases:**
- Document global option sets
- Find option sets for migration
- Reference for development

#### Other Component Types
**Generic Format for:** Attributes (2), Relationships (3, 10), Forms (24), Reports (31), Site Maps (62), Custom Controls (66), Plugin Types (90), Plugin Assemblies (91), SDK Message Steps (92), Canvas Apps (300), Connectors (371), Environment Variables (380, 381), and 80+ other types

**Columns:** Type Code | Component Type | Object ID | Component ID

All components get at least this basic sheet with IDs for reference.

### Workbook Features

**Automatic Sheet Creation**
- Only component types present in the solution get sheets
- If a solution has no flows, there's no flows sheet
- Reduces clutter and focuses on what's actually there

**Professional Formatting**
- Header row with colored background
- Auto-sized columns for readability
- Sorted by Object ID for consistency

**Complete Documentation**
- Every component is documented
- Both in summary (Sheet 2) and detailed sheets
- Perfect for migration planning and audits
- Comparison between solutions
- Documentation summary

---

## Execution Flow

### Step 1: Fetching All Solutions
```
STEP 1: FETCHING ALL SOLUTIONS
======================================================================
  ✓ Location Management (v1.0.0.6)
  ✓ Alerts (v2.0.0.4)
  ✓ Asset Tracking (v1.2.1.0)
  ...

Loaded 25 solutions into memory
```

**What's happening:**
- Queries all visible solutions
- Filters out Microsoft base solutions (unless `--include-microsoft`)
- Stores solution metadata in memory

**Time:** ~5-10 seconds

### Step 2: Collecting All Components
```
STEP 2: COLLECTING ALL COMPONENTS
======================================================================
[1/25] Location Management... ✓ 150 components
[2/25] Alerts... ✓ 45 components
[3/25] Asset Tracking... ✓ 89 components
...

Collected 3,847 total components across all solutions
```

**What's happening:**
- Queries solution components for each solution
- Tracks component reuse across solutions
- Stores all component IDs in memory

**Time:** ~20-30 seconds (for 25 solutions)

### Step 3: Analyzing All Tables
```
STEP 3: ANALYZING ALL TABLES
======================================================================
Found 156 unique tables across all solutions

Fetching metadata and record counts...
  Processed 10/156 tables...
  Processed 20/156 tables...
  ...

✓ Analyzed 156 tables
```

**What's happening:**
- Identifies unique tables across all solutions
- Fetches detailed metadata for each table
- **Uses FetchXML aggregate for unlimited record counts**
- Caches all data in memory

**Time:** ~2-5 minutes (depends on table count and data volume)

### Step 4: Finding Unmanaged Components
```
STEP 4: FINDING UNMANAGED COMPONENTS
======================================================================
  Total components in solutions: 3,847

  Finding unmanaged tables... ✓ Found 15 unmanaged tables
  Finding unmanaged flows... ✓ Found 23 unmanaged flows
  Finding unmanaged canvas apps... ✓ Found 8 unmanaged canvas apps
  Finding unmanaged web resources... ✓ Found 47 unmanaged web resources
  Finding unmanaged forms... ✓ Found 12 unmanaged forms
  Finding unmanaged option sets... ✓ Found 5 unmanaged option sets

  Total unmanaged components: 110
```

**What's happening:**
- Creates a lookup set of all component IDs in solutions
- Queries each component type (tables, flows, apps, web resources, forms, option sets)
- Checks each component against the solution lookup
- Identifies components that are **NOT** in any solution
- These are "orphaned" or "loose" components

**Why This Matters:**
- Unmanaged components won't be exported with solutions
- Helps identify cleanup targets
- Essential for ALM compliance
- Migration readiness check

**Time:** ~30-60 seconds

### Step 5: Creating Master Workbook
```
STEP 5: CREATING MASTER WORKBOOK
======================================================================
  Creating All Solutions Summary...
  Creating All Tables Master List...
  Creating Component Reuse Analysis...
  Creating Table Ownership Map...
  Creating Statistics...
  Creating Unmanaged Tables sheet...
  Creating Unmanaged Flows sheet...
  Creating Unmanaged Canvas Apps sheet...
  Creating Unmanaged Web Resources sheet...
  Creating Unmanaged Forms sheet...
  Creating Unmanaged Option Sets sheet...

✓ Master workbook created: MASTER_Analysis_20260112.xlsx
```

**What's happening:**
- Generates 5 core analysis sheets
- Adds unmanaged component sheets (RED headers) if any found
- All data already in memory (fast)
- Applies formatting and auto-sizing

**Time:** ~10-20 seconds

### Step 6: Creating Individual Workbooks
```
STEP 6: CREATING INDIVIDUAL SOLUTION WORKBOOKS
======================================================================
[1/25] Location Management...
      Fetching component names... ✓
✓ Location Management/Location Management_2024-10-20.xlsx
[2/25] Alerts...
      Fetching component names... ✓
✓ Alerts/Alerts_2024-11-15.xlsx
[3/25] Asset Tracking...
      Fetching component names... ✓
✓ Asset Tracking/Asset Tracking_2024-09-05.xlsx
...

✓ Created 25 individual solution workbooks
```

**What's happening:**
- Creates one folder per solution
- Creates one workbook per solution with:
  - **Summary sheet** with solution metadata
  - **All Components Summary sheet** with logical names and schema names
  - **Component-specific sheets** for each type found (Tables, Flows, Views, Roles, etc.)
- **Fetches logical names and schema names** for all components via API
- **Fetches detailed metadata** for major component types (flows, views, roles, web resources, forms, option sets)
- Reuses cached table data (no additional API calls for tables)
- Applies professional formatting and auto-sizing

**Time:** ~1-3 minutes per solution (depends on component count and types)
- Component name fetching: ~5 seconds per component type
- Detailed metadata fetching: ~10 seconds per component type
- Short timeouts (5 seconds) prevent blocking
- Failed lookups leave fields blank and continue

**Sheet Creation Logic:**
- Always creates: Summary + All Components Summary
- Dynamically creates: One sheet per component type present
- For example, a solution with tables, flows, and views gets 5 total sheets:
  1. Summary
  2. All Components Summary
  3. 1-Entity/Table (detailed table info)
  4. 29-Workflow/Flow (detailed flow info)
  5. 26-Saved Query/View (detailed view info)

### Final Output
```
======================================================================
✓ ANALYSIS COMPLETE
  Environment: contoso
  Extract Date: 20260112
  Master workbook: MASTER_Analysis_20260112.xlsx
  Individual workbooks: 25
  Output folder: /path/to/solution_master_export/Environment Solution Analysis contoso 20260112
======================================================================
```

**Total Time:** ~3-7 minutes (varies by environment size)

---

## Performance & Limitations

### Performance Characteristics

**Small Environment** (5-10 solutions, 50 tables, ~500 components)
- Runtime: ~3-6 minutes
  - Steps 1-3 (Data Collection): ~2-3 minutes
  - Step 4 (Unmanaged Components): ~30 seconds
  - Step 5 (Master Workbook): ~10-15 seconds
  - Step 6 (Individual Workbooks): ~1-2 minutes
- Memory: ~50-100 MB

**Medium Environment** (20-30 solutions, 150 tables, ~3,000 components)
- Runtime: ~9-13 minutes
  - Steps 1-3 (Data Collection): ~5-7 minutes
  - Step 4 (Unmanaged Components): ~45-60 seconds
  - Step 5 (Master Workbook): ~15-20 seconds
  - Step 6 (Individual Workbooks): ~3-5 minutes
- Memory: ~150-250 MB

**Large Environment** (50+ solutions, 300+ tables, ~8,000 components)
- Runtime: ~16-26 minutes
  - Steps 1-3 (Data Collection): ~10-15 minutes
  - Step 4 (Unmanaged Components): ~60-90 seconds
  - Step 5 (Master Workbook): ~20-30 seconds
  - Step 6 (Individual Workbooks): ~5-10 minutes
- Memory: ~500 MB - 1 GB

**Performance Notes:**
- Component name fetching adds ~1-3 minutes per solution in Step 6
- Unmanaged component detection (Step 4) is very fast due to in-memory lookup
- Caching reduces duplicate API calls significantly
- Network speed impacts API call times
- Component metadata (flows, views, roles) takes longer than simple IDs

### Known Limitations

1. **Record Count Timeout**
   - FetchXML aggregate has 15-second timeout per table
   - Very large tables (10M+ records) may timeout
   - Result: Record count shows as "N/A" but script continues

2. **Component Name Fetching Timeout**
   - Each component name lookup has 5-second timeout
   - Complex metadata queries may timeout
   - Result: Logical Name and Schema Name fields left blank, script continues
   - Most common with: Attributes, Relationships, rarely-used component types

3. **Component Limit**
   - Maximum 5,000 components per solution per query
   - Solutions with 5,000+ components may need multiple queries
   - Current version: Single query (handles most cases)

4. **Excel File Size**
   - Solutions with detailed component sheets produce larger files
   - Typical: 500 KB - 5 MB per solution
   - Large solutions with many component types: Up to 20-30 MB

5. **Network Dependency**
   - Requires stable internet connection throughout execution
   - API rate limits may apply (rare for typical usage)
   - Transient failures handled gracefully
   - Component name fetching makes many API calls (one per component type)

### Optimization Tips

**For Faster Execution:**
- Run during off-peak hours
- Use wired connection (not WiFi)
- Close unnecessary applications

**For Large Environments:**
- Consider excluding Microsoft solutions (`--include-microsoft` not used)
- Run analysis on smaller subsets if needed
- Increase timeout values in code if needed

---

## Common Use Cases

### 1. Pre-Migration Analysis
**Goal:** Understand environment before migration

**Steps:**
1. Export from source environment
2. Review Master > Statistics sheet
3. **Review Unmanaged Components sheets** (items that won't migrate with solutions)
4. Review Master > All Tables for data volumes
5. Identify high-impact tables/solutions
6. Plan migration phases by data volume

**Key Metrics:**
- Total Records (overall data volume)
- Tables by solution (migration chunks)
- Component Reuse (dependencies to preserve)
- **Unmanaged Components (manual migration required)**

**Critical Check:**
- Review all Unmanaged component sheets
- Decide: Add to solutions or migrate manually?
- Document manual migration steps

### 2. ALM Compliance & Cleanup
**Goal:** Ensure everything is properly solutioned

**Steps:**
1. Export current environment
2. Open Master > Statistics
3. Review "Unmanaged Components" section
4. For each unmanaged component type:
   - Open the specific sheet (e.g., Unmanaged Tables)
   - Identify owner/creator
   - Decide: Add to solution, delete, or document exception
5. Create action plan for each unmanaged component

**Key Sheets:**
- Unmanaged Tables (orphaned custom entities)
- Unmanaged Flows (automation not in solutions)
- Unmanaged Web Resources (loose script files)
- Unmanaged Forms (orphaned forms)

**Common Actions:**
- Add to existing solution
- Create new solution for component
- Delete if no longer needed
- Document as known exception

### 3. Solution Consolidation
**Goal:** Reduce number of solutions

**Steps:**
1. Export current state
2. Review Master > Component Reuse
3. Review Master > Table Ownership
4. Identify shared components
5. Plan foundation solution
6. Merge overlapping solutions

**Key Sheets:**
- Component Reuse (find duplicates)
- Table Ownership (avoid conflicts)

### 4. Environment Documentation
**Goal:** Document current state for audit/compliance

**Steps:**
1. Export from each environment (DEV, TEST, PROD)
2. Archive timestamped folders
3. Compare Master workbooks across environments
4. Review Individual workbooks for solution details
5. **Document unmanaged components per environment**

**Key Deliverables:**
- Master Analysis (executive summary)
- Individual workbooks (technical details)
- Timestamped archives (version history)
- **Unmanaged component inventory**

### 5. Data Volume Analysis
**Goal:** Understand data footprint by solution

**Steps:**
1. Export environment
2. Open Master > All Solutions
3. Sort by "Total Records" column (descending)
4. Identify top 5-10 data-heavy solutions
5. Review individual workbooks for table breakdown

**Decision Points:**
- Which solutions to migrate first?
- Which tables need archival before migration?
- Which solutions can be excluded from backups?

### 6. Dependency Discovery
**Goal:** Find cross-solution dependencies

**Steps:**
1. Export environment
2. Review Master > Component Reuse
3. Review Master > Table Ownership
4. Look for tables in 3+ solutions
5. Map dependencies before changes

**Red Flags:**
- Core tables in many solutions (risk of conflicts)
- Components shared across 5+ solutions (coupling)

### 7. Cleanup Planning
**Goal:** Identify unused/empty components

**Steps:**
1. Export environment
2. Review Master > All Tables
3. Filter by Record Count = 0 or N/A
4. Review Master > Statistics > Empty Tables
5. **Review all Unmanaged Components sheets**
6. Identify solutions with empty tables
7. **Identify orphaned components to delete**
8. Plan cleanup/removal

**Cleanup Targets:**
- Empty tables (no data)
- **Unmanaged flows** (test automation never solutioned)
- **Unmanaged web resources** (old JS/CSS files)
- **Unmanaged forms** (backup/test forms)
- Tables with < 10 records (potential test data)

**Questions to Ask:**
- Is this component still used?
- Who created it and when?
- Is it referenced by active solutions?
- **Why isn't it in a solution?**
- Can it be safely deleted?

---

## Troubleshooting

### Issue: "Authentication failed"

**Possible Causes:**
- Insufficient permissions
- Wrong environment URL
- Expired device code

**Solutions:**
- Verify System Administrator or System Customizer role
- Check URL format: `https://orgname.crm.dynamics.com`
- Complete authentication within 15 minutes
- Try different browser if consent screen doesn't load

### Issue: "No solutions found"

**Possible Causes:**
- All solutions are Microsoft base solutions
- No custom solutions exist
- Incorrect environment URL

**Solutions:**
- Use `--include-microsoft` flag to see all solutions
- Verify environment URL is correct
- Check if solutions exist in Power Platform admin center

### Issue: Record counts show "N/A"

**Possible Causes:**
- Table is very large (10M+ records)
- FetchXML timeout (15 seconds)
- Permissions issue on table

**Solutions:**
- Not an error - script continues normally
- Large tables may legitimately timeout
- Check individual table permissions if many show N/A

### Issue: Script runs slowly

**Possible Causes:**
- Large environment (many solutions/tables)
- Network latency
- High API usage

**Solutions:**
- Expected for large environments (10-15 min is normal)
- Run during off-peak hours
- Check internet connection stability
- Close other applications

### Issue: Excel file won't open

**Possible Causes:**
- File corruption during save
- Excel version compatibility
- Disk space issue

**Solutions:**
- Re-run export (rare issue)
- Use Excel 2016 or later
- Ensure sufficient disk space (100+ MB free)

### Issue: Missing components in individual workbooks

**Possible Causes:**
- Solution has 5,000+ components (API limit)
- Solution was recently modified
- Component metadata unavailable

**Solutions:**
- Check "Total Components" in Summary sheet
- Component may be in "All Components" sheet but not detailed sheets
- Re-run export if environment was recently changed

---

## Best Practices

### Before Running

1. **Verify Access**
   - Confirm you have System Administrator or System Customizer role
   - Test access by opening environment in browser

2. **Check Environment State**
   - Avoid running during active deployments
   - Run after hours for large environments
   - Ensure environment is stable

3. **Plan Output Location**
   - Ensure sufficient disk space (100+ MB per environment)
   - Use descriptive output directory names
   - Create backup if overwriting existing exports

### During Execution

1. **Don't Interrupt**
   - Let script run to completion
   - Don't close terminal/command prompt
   - Don't put computer to sleep

2. **Monitor Progress**
   - Watch console output for errors
   - Note any "N/A" record counts
   - Check for authentication timeouts

3. **Network Stability**
   - Use wired connection if possible
   - Avoid VPN if causing latency
   - Don't start large downloads during run

### After Completion

1. **Verify Output**
   - Check that master workbook opens
   - Spot-check 2-3 individual workbooks
   - Verify record counts seem reasonable

2. **Archive Properly**
   - Keep timestamped folders intact
   - Don't rename generated files
   - Add README if needed for future reference

3. **Share Appropriately**
   - Master workbook for executives/managers
   - Individual workbooks for technical teams
   - Be mindful of sensitive data in exports

### Regular Execution

1. **Weekly/Monthly Exports**
   - Track environment changes over time
   - Compare master workbooks to see trends
   - Document major changes

2. **Pre/Post Deployment**
   - Export before major deployments
   - Export after to verify changes
   - Compare master workbooks

3. **Quarterly Reviews**
   - Use for environment health checks
   - Review Statistics sheet for growth
   - Plan cleanup based on trends

---

## Advanced Usage

### Custom Output Directory Structure

Create organized export structure:
```bash
# Create structure
mkdir -p PowerPlatform/Exports/{DEV,TEST,PROD}

# Export to each
python3 solution_exporter_with_master.py \
  --url https://contoso-dev.crm.dynamics.com \
  --output-dir PowerPlatform/Exports/DEV

python3 solution_exporter_with_master.py \
  --url https://contoso-test.crm.dynamics.com \
  --output-dir PowerPlatform/Exports/TEST

python3 solution_exporter_with_master.py \
  --url https://contoso.crm.dynamics.com \
  --output-dir PowerPlatform/Exports/PROD
```

Result:
```
PowerPlatform/Exports/
├── DEV/
│   └── Environment Solution Analysis contoso-dev 20260112/
├── TEST/
│   └── Environment Solution Analysis contoso-test 20260112/
└── PROD/
    └── Environment Solution Analysis contoso 20260112/
```

### Scheduled Execution

**Windows Task Scheduler:**
```batch
@echo off
cd C:\PowerPlatform\Scripts
python solution_exporter_with_master.py --url https://contoso.crm.dynamics.com --output-dir C:\PowerPlatform\Exports
```

**Linux Cron:**
```bash
0 2 * * 0 cd /home/user/pp-scripts && python3 solution_exporter_with_master.py --url https://contoso.crm.dynamics.com --output-dir /home/user/pp-exports
```

**Note:** Scheduled execution requires cached authentication or service principal (custom implementation needed).

### Comparing Environments

Export from multiple environments, then compare:

1. **Export all environments**
2. **Open master workbooks side-by-side**
3. **Compare Statistics sheets**
4. **Compare All Solutions sheets**
5. **Identify differences**

**Key Comparisons:**
- Solution versions (DEV ahead of PROD?)
- Component counts (drift detection)
- Record counts (data sync verification)
- Publisher differences

---

## FAQ

**Q: How long does it take to run?**
A: 3-7 minutes for typical environments (20-30 solutions, 150 tables). Large environments may take 10-15 minutes.

**Q: Can I run this on multiple environments simultaneously?**
A: Yes, each run is independent. Open multiple terminal windows.

**Q: Will this modify my environment?**
A: No, this tool is read-only. It only queries data, never writes.

**Q: Can I automate this?**
A: Authentication requires interactive browser login. For automation, implement service principal authentication (custom code needed).

**Q: What if I have 100+ solutions?**
A: Script will handle it, but may take 15-20 minutes. Consider using `--include-microsoft` flag to see true count.

**Q: Why do some tables show "N/A" for record count?**
A: Very large tables (10M+ records) may timeout during count query. Not an error - script continues normally.

**Q: Can I export only specific solutions?**
A: Not directly. Export all, then delete unwanted individual workbooks. Master workbook will still show all solutions.

**Q: How often should I run this?**
A: Depends on environment change frequency:
- Active development: Weekly
- Stable production: Monthly
- Before/after major changes: Always

**Q: Can I share the output files?**
A: Yes, but be mindful that they contain metadata about your environment. Review for sensitive information before sharing externally.

**Q: What Excel version do I need?**
A: Excel 2016 or later recommended. Files are standard .xlsx format.

**Q: Will this work with Dynamics 365 apps?**
A: Yes, this works with any Power Platform / Dataverse environment, including those with Dynamics 365 apps.

**Q: Can I cancel mid-execution?**
A: Yes (Ctrl+C), but you'll need to restart from the beginning. No partial results are saved.

---

## Support & Feedback

### Getting Help

1. **Check This Guide** - Most common issues are covered
2. **Review Console Output** - Error messages are descriptive
3. **Verify Prerequisites** - Ensure all packages installed
4. **Test Permissions** - Confirm environment access

### Reporting Issues

When reporting issues, include:
- Exact command used
- Console output (full text)
- Environment size (approx. # of solutions)
- Python version (`python --version`)
- Operating system

### Feature Requests

Common requests (not currently implemented):
- Service principal authentication
- Solution filtering
- Custom report formats
- API-based scheduling
- Differential analysis (compare two exports)

---

## Version History

**Current Version:** 2.0 (Master Analysis Version)

**Features:**
- Master workbook with cross-solution analysis
- Individual solution workbooks
- In-memory data collection
- Component reuse tracking
- Table ownership mapping
- Record count aggregation (unlimited via FetchXML)
- Environment-timestamped output folders
- Solution-specific subfolders

---

## Practical Examples: Unmanaged Components

### Example 1: Pre-Migration Cleanup

**Scenario:** Migrating environment from DEV to PROD

**Steps:**
1. Run analysis: `python3 solution_exporter_with_master.py --url https://dev.crm.dynamics.com`
2. Open `MASTER_Analysis_{date}.xlsx`
3. Check Statistics sheet - shows 47 unmanaged components
4. Review each unmanaged component sheet:

**Unmanaged Tables Sheet:**
```
ContactHistory     | cr123_contacthistory  | 1,523 records
TempCalculations   | cr123_tempcalc        | 0 records
TestData           | cr123_testdata        | 45 records
```

**Decision:**
- ContactHistory → Add to "Core Extensions" solution
- TempCalculations → Delete (empty, no longer used)
- TestData → Delete (test data not for PROD)

**Unmanaged Web Resources Sheet:**
```
dev_testscript.js  | JavaScript  | 2024-03-15
old_styles.css     | CSS         | 2023-11-20
```

**Decision:**
- Both → Delete (leftover from old development)

**Result:** Clean environment ready for migration

### Example 2: ALM Compliance Audit

**Scenario:** Quarterly compliance check

**Steps:**
1. Run analysis on PROD environment
2. Open Statistics sheet
3. Section shows:
   - Unmanaged Tables: 3
   - Unmanaged Flows: 8
   - Unmanaged Canvas Apps: 2
   - Total Unmanaged: 13

**Investigation:**
Open each unmanaged sheet, note GUIDs, research in Power Platform:
- 3 tables created by users directly (not through solutions)
- 8 flows created in maker portal (personal automation)
- 2 canvas apps in development (not production-ready)

**Action Plan:**
- Email users to add components to solutions
- Document deadline for compliance
- Schedule follow-up audit in 30 days

### Example 3: Environment Cleanup Project

**Scenario:** Reduce environment bloat

**Steps:**
1. Run analysis
2. Check Statistics:
   - Unmanaged Web Resources: 127 (!!)
   - Unmanaged Forms: 34
   
**Investigation:**
Open Unmanaged Web Resources sheet, filter by Created On:
- 85 files from 2022-2023 (likely obsolete)
- 42 files from 2024 (investigate)

**Action:**
- Research 2024 files → 15 still in use, add to solutions
- Delete all 2022-2023 files after backup
- Document deleted resources

**Result:** Reduced from 127 → 15 unmanaged web resources

### Example 4: Finding Component Owners

**Scenario:** Who created these unmanaged components?

**Steps:**
1. Run analysis, find 23 unmanaged flows
2. Note the Workflow IDs from Unmanaged Flows sheet
3. Use Power Platform API to query audit data:

```powershell
# Example PowerShell to find creator
$workflowId = "a1b2c3d4-e5f6-7890-1234-567890abcdef"
Get-CrmRecords -EntityLogicalName workflow -FilterAttribute workflowid -FilterOperator eq -FilterValue $workflowId -Fields createdby,createdon
```

**Result:** Identify creators, reach out for solution ownership

### Example 5: Comparing Environments

**Scenario:** Compare DEV vs PROD for unmanaged components

**Steps:**
1. Run analysis on both environments
2. Compare Statistics sheets:

**DEV Environment:**
- Unmanaged Tables: 15
- Unmanaged Flows: 34
- Total Unmanaged: 89

**PROD Environment:**
- Unmanaged Tables: 3
- Unmanaged Flows: 5
- Total Unmanaged: 11

**Analysis:**
- DEV has appropriate unmanaged components (development/testing)
- PROD is clean (expected for production)
- Those 11 unmanaged components in PROD need investigation

**Action:** Review PROD unmanaged components, add to solutions or delete

---

## Quick Reference

### Installation
```bash
pip install msal requests openpyxl --break-system-packages
```

### Basic Command
```bash
python3 solution_exporter_with_master.py --url https://yourorg.crm.dynamics.com
```

### Output Location
```
{output-dir}/Environment Solution Analysis {env-name} {YYYYMMDD}/
```

### Key Files
- `MASTER_Analysis_{YYYYMMDD}.xlsx` - Cross-solution analysis + Unmanaged components
- `{SolutionName}/{SolutionName}_{date}.xlsx` - Individual solution details

### Master Workbook Sheets
- All Solutions (component counts + record totals)
- All Tables (record counts for all tables)
- Component Reuse (shared components)
- Table Ownership (tables by solution)
- Statistics (environment metrics + **unmanaged counts**)
- **Unmanaged Tables** (RED header - custom tables not in solutions)
- **Unmanaged Flows** (RED header - flows not in solutions)
- **Unmanaged Canvas Apps** (RED header - apps not in solutions)
- **Unmanaged Web Resources** (RED header - files not in solutions)
- **Unmanaged Forms** (RED header - forms not in solutions)
- **Unmanaged Option Sets** (RED header - option sets not in solutions)

### Typical Runtime
- Small: 3-6 minutes
- Medium: 9-13 minutes  
- Large: 16-26 minutes

### Authentication
- Device code flow (interactive)
- Browser-based sign-in
- System Admin or System Customizer role required

---

**End of Usage Guide**
