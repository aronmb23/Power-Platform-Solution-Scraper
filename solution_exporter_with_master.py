#!/usr/bin/env python3
"""
Power Platform Solution Component Exporter - Master Analysis Version

This performs the following action:
1. Collects ALL solution data in memory first
2. Performs cross-solution analysis
3. Creates individual solution workbooks
4. Creates a MASTER workbook with:
   - All solutions summary
   - All tables across all solutions with record counts
   - Component reuse analysis (which components appear in multiple solutions)
   - Table ownership map (which solution "owns" each table)
   - Dependency analysis
   - Comparison sheets

Usage:
    python3 solution_exporter_with_master.py --url https://yourorg.crm.dynamics.com
"""

import requests
import json
import argparse
from pathlib import Path
from datetime import datetime
import sys
from collections import defaultdict
import xml.etree.ElementTree as ET
from urllib.parse import quote

try:
    from msal import PublicClientApplication
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
except ImportError:
    print("ERROR: Required packages not found.")
    print("Install with: pip install msal requests openpyxl")
    sys.exit(1)


class MasterSolutionAnalyzer:
    """Analyzes all solutions and creates master comparison workbook"""
    
    COMPONENT_TYPES = {
        1: 'Entity/Table',
        2: 'Attribute/Column',
        3: 'Relationship',
        9: 'Option Set',
        10: 'Entity Relationship',
        20: 'Role',
        24: 'Form',
        26: 'Saved Query/View',
        29: 'Workflow/Flow',
        31: 'Report',
        60: 'System Form',
        61: 'Web Resource',
        62: 'Site Map',
        66: 'Custom Control',
        90: 'Plugin Type',
        91: 'Plugin Assembly',
        92: 'SDK Message Processing Step',
        300: 'Canvas App',
        371: 'Connector',
        380: 'Environment Variable Definition',
        381: 'Environment Variable Value',
        10293: 'View'
    }
    
    def __init__(self, environment_url, access_token=None):
        self.environment_url = environment_url.rstrip('/')
        self.api_url = f"{self.environment_url}/api/data/v9.2"
        self.access_token = access_token
        
        # Extract environment name from URL
        # URL format: https://orgname.crm.dynamics.com or https://orgname.crm4.dynamics.com
        self.environment_name = self._extract_environment_name(environment_url)
        
        # In-memory storage for ALL data
        self.all_solutions = []  # List of all solution info
        self.all_components = {}  # solution_id -> list of components
        self.all_tables = {}  # table_logical_name -> full metadata
        self.all_flows = {}  # flow_id -> metadata
        self.all_views = {}  # view_id -> metadata
        self.all_roles = {}  # role_id -> metadata
        self.all_web_resources = {}  # wr_id -> metadata
        self.all_relationships = defaultdict(list)  # table_name -> list of relationships
        
        # Analysis data structures
        self.table_to_solutions = defaultdict(set)  # table_name -> set of solution names
        self.component_reuse = defaultdict(list)  # component_id -> list of solution names
        self.table_record_counts = {}  # table_name -> count
        
        # Caches
        self.component_cache = {}
        self.record_count_cache = {}
    
    def _extract_environment_name(self, url):
        """Extract environment name from Dataverse URL"""
        # Remove protocol
        url = url.replace('https://', '').replace('http://', '')
        
        # Get the first part (orgname.crm.dynamics.com -> orgname)
        parts = url.split('.')
        if parts:
            env_name = parts[0]
            # Clean up for folder name
            env_name = "".join(c for c in env_name if c.isalnum() or c in ('-', '_'))
            return env_name if env_name else 'Unknown'
        
        return 'Unknown'
    
    def authenticate_interactive(self, client_id=None, tenant_id=None):
        """Authenticate using device code flow"""
        if not client_id:
            client_id = "51f81489-12ee-4a9e-aaae-a2591f45987d"
        
        if not tenant_id:
            tenant_id = "common"
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = PublicClientApplication(client_id=client_id, authority=authority)
        scopes = [f"{self.environment_url}/.default"]
        
        print("\n" + "="*70)
        print("AUTHENTICATION REQUIRED")
        print("="*70)
        
        flow = app.initiate_device_flow(scopes=scopes)
        
        if "user_code" not in flow:
            raise Exception(f"Failed to create device flow")
        
        print(flow["message"])
        print("="*70 + "\n")
        
        result = app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✓ Authentication successful!\n")
            return True
        else:
            print(f"✗ Authentication failed")
            return False
    
    def get_headers(self):
        return {
            'Authorization': f'Bearer {self.access_token}',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
            'Content-Type': 'application/json; charset=utf-8',
            'Prefer': 'odata.include-annotations="*"'
        }
    
    def get_record_count_fetchxml(self, entity_name, primary_id):
        """Get record count using FetchXML aggregate"""
        fetchxml = f"""
        <fetch aggregate="true">
            <entity name="{entity_name}">
                <attribute name="{primary_id}" alias="count" aggregate="count" />
            </entity>
        </fetch>
        """
        
        try:
            encoded_fetch = quote(fetchxml)
            query = f"{self.api_url}/{entity_name}s?fetchXml={encoded_fetch}"
            
            response = requests.get(query, headers=self.get_headers(), timeout=15)
            
            if response.status_code == 200:
                data = response.json()
                values = data.get('value', [])
                
                if values and len(values) > 0:
                    count = values[0].get('count', 0)
                    if isinstance(count, str):
                        count = int(count) if count.isdigit() else 0
                    return count
            
            return None
        except:
            return None
    
    def get_component_logical_name(self, component_type, object_id):
        """Get logical name and schema name for a component based on its type"""
        logical_name = ''
        schema_name = ''
        
        try:
            # Tables - already have this data
            if component_type == 1:
                for tbl in self.all_tables.values():
                    if tbl.get('ObjectId') == object_id:
                        logical_name = tbl['LogicalName']
                        schema_name = tbl['SchemaName']
                        break
            
            # Workflows/Flows
            elif component_type == 29:
                query = f"{self.api_url}/workflows({object_id})"
                params = {'$select': 'uniquename,name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    flow = response.json()
                    logical_name = flow.get('uniquename', '')
                    schema_name = flow.get('name', '')
            
            # Views/Saved Queries
            elif component_type in [26, 10293]:
                query = f"{self.api_url}/savedqueries({object_id})"
                params = {'$select': 'name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    view = response.json()
                    logical_name = view.get('name', '')
                    schema_name = view.get('name', '')
            
            # Security Roles
            elif component_type == 20:
                query = f"{self.api_url}/roles({object_id})"
                params = {'$select': 'name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    role = response.json()
                    logical_name = role.get('name', '')
                    schema_name = role.get('name', '')
            
            # Web Resources
            elif component_type == 61:
                query = f"{self.api_url}/webresources({object_id})"
                params = {'$select': 'name,displayname'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    wr = response.json()
                    logical_name = wr.get('name', '')
                    schema_name = wr.get('name', '')
            
            # Forms
            elif component_type in [24, 60]:
                query = f"{self.api_url}/systemforms({object_id})"
                params = {'$select': 'name,uniquename'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    form = response.json()
                    logical_name = form.get('uniquename', form.get('name', ''))
                    schema_name = form.get('name', '')
            
            # Option Sets
            elif component_type == 9:
                query = f"{self.api_url}/GlobalOptionSetDefinitions({object_id})"
                params = {'$select': 'Name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    optionset = response.json()
                    logical_name = optionset.get('Name', '')
                    schema_name = optionset.get('Name', '')
            
            # Plugin Assemblies
            elif component_type == 91:
                query = f"{self.api_url}/pluginassemblies({object_id})"
                params = {'$select': 'name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    assembly = response.json()
                    logical_name = assembly.get('name', '')
                    schema_name = assembly.get('name', '')
            
            # Plugin Types
            elif component_type == 90:
                query = f"{self.api_url}/plugintypes({object_id})"
                params = {'$select': 'typename,name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    plugin = response.json()
                    logical_name = plugin.get('typename', '')
                    schema_name = plugin.get('name', '')
            
            # SDK Message Processing Steps
            elif component_type == 92:
                query = f"{self.api_url}/sdkmessageprocessingsteps({object_id})"
                params = {'$select': 'name'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    step = response.json()
                    logical_name = step.get('name', '')
                    schema_name = step.get('name', '')
            
            # Canvas Apps
            elif component_type == 300:
                query = f"{self.api_url}/canvasapps({object_id})"
                params = {'$select': 'name,displayname'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    app = response.json()
                    logical_name = app.get('name', '')
                    schema_name = app.get('displayname', app.get('name', ''))
            
            # Environment Variable Definitions
            elif component_type == 380:
                query = f"{self.api_url}/environmentvariabledefinitions({object_id})"
                params = {'$select': 'schemaname,displayname'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    env_var = response.json()
                    logical_name = env_var.get('schemaname', '')
                    schema_name = env_var.get('displayname', '')
            
            # Relationships
            elif component_type in [3, 10]:
                query = f"{self.api_url}/RelationshipDefinitions({object_id})"
                params = {'$select': 'SchemaName'}
                response = requests.get(query, headers=self.get_headers(), params=params, timeout=5)
                if response.status_code == 200:
                    rel = response.json()
                    logical_name = rel.get('SchemaName', '')
                    schema_name = rel.get('SchemaName', '')
        
        except:
            pass
        
        return logical_name, schema_name
    
    def fetch_all_solutions(self, include_microsoft=False):
        """Fetch all solutions and store in memory"""
        print("="*70)
        print("STEP 1: FETCHING ALL SOLUTIONS")
        print("="*70 + "\n")
        
        query = f"{self.api_url}/solutions"
        params = {
            '$select': 'solutionid,uniquename,friendlyname,version,ismanaged,installedon,description,publisherid',
            '$expand': 'publisherid($select=friendlyname,uniquename)',
            '$orderby': 'installedon asc',
            '$filter': "isvisible eq true"
        }
        
        try:
            response = requests.get(query, headers=self.get_headers(), params=params)
            response.raise_for_status()
            all_solutions = response.json().get('value', [])
            
            for solution in all_solutions:
                unique_name = solution['uniquename']
                
                if not include_microsoft:
                    if unique_name.startswith(('Default', 'Crd0a29', 'msdyn_', 'Dynamics', 'Microsoft', 'msdynce_', 'HCMCommon')):
                        continue
                
                solution_info = {
                    'solution_id': solution['solutionid'],
                    'unique_name': unique_name,
                    'display_name': solution['friendlyname'],
                    'version': solution['version'],
                    'publisher': solution.get('publisherid', {}).get('friendlyname', 'Unknown'),
                    'publisher_unique': solution.get('publisherid', {}).get('uniquename', 'Unknown'),
                    'is_managed': solution['ismanaged'],
                    'installed_on': solution.get('installedon'),
                    'description': solution.get('description', '')
                }
                
                self.all_solutions.append(solution_info)
                print(f"  ✓ {solution['friendlyname']} (v{solution['version']})")
            
            print(f"\nLoaded {len(self.all_solutions)} solutions into memory\n")
            return len(self.all_solutions)
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching solutions: {e}")
            return 0
    
    def collect_all_components(self):
        """Collect ALL components from ALL solutions into memory"""
        print("="*70)
        print("STEP 2: COLLECTING ALL COMPONENTS")
        print("="*70 + "\n")
        
        for i, solution in enumerate(self.all_solutions, 1):
            print(f"[{i}/{len(self.all_solutions)}] {solution['display_name']}...", end=' ', flush=True)
            
            # Get components
            query = f"{self.api_url}/solutioncomponents"
            params = {
                '$select': 'objectid,componenttype,solutioncomponentid',
                '$filter': f"_solutionid_value eq {solution['solution_id']}",
                '$top': 5000
            }
            
            try:
                response = requests.get(query, headers=self.get_headers(), params=params)
                response.raise_for_status()
                components = response.json().get('value', [])
                
                self.all_components[solution['solution_id']] = components
                
                # Track component reuse
                for comp in components:
                    comp_key = f"{comp['componenttype']}_{comp['objectid']}"
                    self.component_reuse[comp_key].append(solution['display_name'])
                
                print(f"✓ {len(components)} components")
                
            except:
                print(f"✗ Error")
                self.all_components[solution['solution_id']] = []
        
        total_components = sum(len(comps) for comps in self.all_components.values())
        print(f"\nCollected {total_components} total components across all solutions\n")
    
    def analyze_all_tables(self):
        """Analyze all tables across all solutions"""
        print("="*70)
        print("STEP 3: ANALYZING ALL TABLES")
        print("="*70 + "\n")
        
        all_table_ids = set()
        
        # Collect unique table IDs
        for solution_id, components in self.all_components.items():
            for comp in components:
                if comp['componenttype'] == 1:  # Entity/Table
                    all_table_ids.add(comp['objectid'])
        
        print(f"Found {len(all_table_ids)} unique tables across all solutions\n")
        print("Fetching metadata and record counts...\n")
        
        for i, table_id in enumerate(all_table_ids, 1):
            try:
                query = f"{self.api_url}/EntityDefinitions({table_id})"
                params = {
                    '$select': 'LogicalName,SchemaName,DisplayName,Description,PrimaryIdAttribute,PrimaryNameAttribute,ObjectTypeCode,IsCustomEntity'
                }
                
                response = requests.get(query, headers=self.get_headers(), params=params)
                
                if response.status_code == 200:
                    entity = response.json()
                    
                    display_name = entity.get('DisplayName', {})
                    if isinstance(display_name, dict):
                        display_name = display_name.get('UserLocalizedLabel', {}).get('Label', entity['LogicalName'])
                    
                    logical_name = entity['LogicalName']
                    primary_id = entity.get('PrimaryIdAttribute', '')
                    
                    # Get record count
                    if logical_name not in self.record_count_cache:
                        record_count = self.get_record_count_fetchxml(logical_name, primary_id)
                        self.record_count_cache[logical_name] = record_count
                    else:
                        record_count = self.record_count_cache[logical_name]
                    
                    self.all_tables[logical_name] = {
                        'ObjectId': table_id,  # Store objectid for matching
                        'LogicalName': logical_name,
                        'SchemaName': entity['SchemaName'],
                        'DisplayName': display_name,
                        'Description': entity.get('Description', {}).get('UserLocalizedLabel', {}).get('Label', ''),
                        'PrimaryIdAttribute': primary_id,
                        'PrimaryNameAttribute': entity.get('PrimaryNameAttribute', ''),
                        'ObjectTypeCode': entity.get('ObjectTypeCode', ''),
                        'IsCustomEntity': entity.get('IsCustomEntity', False),
                        'RecordCount': record_count
                    }
                    
                    # Cache the table_id -> logical_name mapping for quick lookup
                    self.component_cache[str(table_id)] = {'LogicalName': logical_name}
                    
                    # Track which solutions contain this table
                    for solution in self.all_solutions:
                        solution_components = self.all_components.get(solution['solution_id'], [])
                        for comp in solution_components:
                            if comp['componenttype'] == 1 and comp['objectid'] == table_id:
                                self.table_to_solutions[logical_name].add(solution['display_name'])
                    
                    if i % 10 == 0:
                        print(f"  Processed {i}/{len(all_table_ids)} tables...")
                
            except:
                pass
        
        print(f"\n✓ Analyzed {len(self.all_tables)} tables\n")
    
    def create_master_workbook(self, output_dir):
        """Create master analysis workbook"""
        print("="*70)
        print("STEP 4: CREATING MASTER WORKBOOK")
        print("="*70 + "\n")
        
        filepath = output_dir / f"MASTER_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        # Styles
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        # SHEET 1: All Solutions Summary
        print("  Creating All Solutions Summary...")
        ws_solutions = wb.create_sheet("All Solutions")
        ws_solutions.append(['Solution Name', 'Unique Name', 'Version', 'Publisher', 'Type', 'Installed On', 'Total Components', 'Tables', 'Flows', 'Views', 'Roles', 'Total Records'])
        
        for row in ws_solutions['A1:L1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        for solution in sorted(self.all_solutions, key=lambda s: s['display_name']):
            components = self.all_components.get(solution['solution_id'], [])
            
            table_count = sum(1 for c in components if c['componenttype'] == 1)
            flow_count = sum(1 for c in components if c['componenttype'] == 29)
            view_count = sum(1 for c in components if c['componenttype'] in [26, 10293])
            role_count = sum(1 for c in components if c['componenttype'] == 20)
            
            # Calculate total record count for tables in this solution
            total_records = 0
            for comp in components:
                if comp['componenttype'] == 1:  # Table component
                    # Find matching table by objectid
                    for table_info in self.all_tables.values():
                        if table_info.get('ObjectId') == comp['objectid']:
                            if table_info.get('RecordCount') is not None:
                                total_records += table_info['RecordCount']
                            break
            
            total_records_display = f"{total_records:,}" if total_records > 0 else "0"
            
            ws_solutions.append([
                solution['display_name'],
                solution['unique_name'],
                solution['version'],
                solution['publisher'],
                'Managed' if solution['is_managed'] else 'Unmanaged',
                solution['installed_on'][:10] if solution['installed_on'] else '',
                len(components),
                table_count,
                flow_count,
                view_count,
                role_count,
                total_records_display
            ])
        
        for column in ws_solutions.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws_solutions.column_dimensions[column_letter].width = min(max_length + 2, 50)
        
        # SHEET 2: All Tables Master List
        print("  Creating All Tables Master List...")
        ws_tables = wb.create_sheet("All Tables")
        ws_tables.append(['Table Name', 'Logical Name', 'Schema Name', 'Record Count', 'Is Custom', 'In # Solutions', 'Solutions'])
        
        for row in ws_tables['A1:G1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        for table_name, table_info in sorted(self.all_tables.items(), key=lambda x: x[1]['DisplayName']):
            solutions_with_table = self.table_to_solutions.get(table_name, set())
            
            record_count_display = f"{table_info['RecordCount']:,}" if table_info['RecordCount'] is not None else "N/A"
            
            ws_tables.append([
                table_info['DisplayName'],
                table_info['LogicalName'],
                table_info['SchemaName'],
                record_count_display,
                'Yes' if table_info['IsCustomEntity'] else 'No',
                len(solutions_with_table),
                ', '.join(sorted(solutions_with_table))
            ])
        
        for column in ws_tables.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws_tables.column_dimensions[column_letter].width = min(max_length + 2, 60)
        
        # SHEET 3: Component Reuse Analysis
        print("  Creating Component Reuse Analysis...")
        ws_reuse = wb.create_sheet("Component Reuse")
        ws_reuse.append(['Component Type', 'Component ID', 'Used In # Solutions', 'Solutions'])
        
        for row in ws_reuse['A1:D1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        # Find components used in multiple solutions
        reused_components = {k: v for k, v in self.component_reuse.items() if len(v) > 1}
        
        for comp_key, solution_list in sorted(reused_components.items(), key=lambda x: len(x[1]), reverse=True):
            comp_type_num, comp_id = comp_key.split('_', 1)
            comp_type = self.COMPONENT_TYPES.get(int(comp_type_num), f"Type {comp_type_num}")
            
            ws_reuse.append([
                comp_type,
                comp_id,
                len(solution_list),
                ', '.join(sorted(solution_list))
            ])
        
        for column in ws_reuse.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws_reuse.column_dimensions[column_letter].width = min(max_length + 2, 80)
        
        # SHEET 4: Table Ownership Map
        print("  Creating Table Ownership Map...")
        ws_ownership = wb.create_sheet("Table Ownership")
        ws_ownership.append(['Table Name', 'Logical Name', 'Record Count', 'Primary Owner Solution', '# Solutions', 'All Solutions'])
        
        for row in ws_ownership['A1:F1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        for table_name, table_info in sorted(self.all_tables.items(), key=lambda x: x[1]['DisplayName']):
            solutions_with_table = list(self.table_to_solutions.get(table_name, set()))
            
            # Determine "primary owner" - first solution alphabetically (or oldest)
            primary_owner = solutions_with_table[0] if solutions_with_table else 'Unknown'
            
            record_count_display = f"{table_info['RecordCount']:,}" if table_info['RecordCount'] is not None else "N/A"
            
            ws_ownership.append([
                table_info['DisplayName'],
                table_info['LogicalName'],
                record_count_display,
                primary_owner,
                len(solutions_with_table),
                ', '.join(sorted(solutions_with_table))
            ])
        
        for column in ws_ownership.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws_ownership.column_dimensions[column_letter].width = min(max_length + 2, 60)
        
        # SHEET 5: Statistics
        print("  Creating Statistics...")
        ws_stats = wb.create_sheet("Statistics")
        ws_stats.append(['Metric', 'Value'])
        
        for row in ws_stats['A1:B1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        total_components = sum(len(comps) for comps in self.all_components.values())
        total_tables = len(self.all_tables)
        custom_tables = sum(1 for t in self.all_tables.values() if t['IsCustomEntity'])
        tables_with_data = sum(1 for t in self.all_tables.values() if t['RecordCount'] and t['RecordCount'] > 0)
        total_records = sum(t['RecordCount'] for t in self.all_tables.values() if t['RecordCount'])
        
        managed_solutions = sum(1 for s in self.all_solutions if s['is_managed'])
        
        ws_stats.append(['Total Solutions', len(self.all_solutions)])
        ws_stats.append(['Managed Solutions', managed_solutions])
        ws_stats.append(['Unmanaged Solutions', len(self.all_solutions) - managed_solutions])
        ws_stats.append(['Total Components', total_components])
        ws_stats.append(['Total Unique Tables', total_tables])
        ws_stats.append(['Custom Tables', custom_tables])
        ws_stats.append(['System Tables', total_tables - custom_tables])
        ws_stats.append(['Tables with Data', tables_with_data])
        ws_stats.append(['Empty Tables', total_tables - tables_with_data])
        ws_stats.append(['Total Records (All Tables)', f"{total_records:,}"])
        ws_stats.append(['Components Reused Across Solutions', len([c for c, s in self.component_reuse.items() if len(s) > 1])])
        
        ws_stats.column_dimensions['A'].width = 40
        ws_stats.column_dimensions['B'].width = 30
        
        # Save
        wb.save(filepath)
        print(f"\n✓ Master workbook created: {filepath.name}\n")
        
        return filepath
    
    def get_flow_metadata(self, flow_id):
        """Get detailed flow metadata"""
        if flow_id in self.all_flows:
            return self.all_flows[flow_id]
        
        try:
            query = f"{self.api_url}/workflows({flow_id})"
            params = {
                '$select': 'workflowid,name,uniquename,description,category,primaryentity,type,statecode,createdon,modifiedon,createdby,modifiedby'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                flow = response.json()
                self.all_flows[flow_id] = flow
                return flow
        except:
            pass
        
        return None
    
    def get_view_metadata(self, view_id):
        """Get detailed view metadata"""
        if view_id in self.all_views:
            return self.all_views[view_id]
        
        try:
            query = f"{self.api_url}/savedqueries({view_id})"
            params = {
                '$select': 'savedqueryid,name,description,returnedtypecode,querytype,isdefault,isquickfindquery,isprivate,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                view = response.json()
                self.all_views[view_id] = view
                return view
        except:
            pass
        
        return None
    
    def get_role_metadata(self, role_id):
        """Get detailed role metadata"""
        if role_id in self.all_roles:
            return self.all_roles[role_id]
        
        try:
            query = f"{self.api_url}/roles({role_id})"
            params = {
                '$select': 'roleid,name,businessunitid,iscustomizable,ismanaged,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                role = response.json()
                self.all_roles[role_id] = role
                return role
        except:
            pass
        
        return None
    
    def get_web_resource_metadata(self, wr_id):
        """Get detailed web resource metadata"""
        if wr_id in self.all_web_resources:
            return self.all_web_resources[wr_id]
        
        try:
            query = f"{self.api_url}/webresources({wr_id})"
            params = {
                '$select': 'webresourceid,name,displayname,description,webresourcetype,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                wr = response.json()
                self.all_web_resources[wr_id] = wr
                return wr
        except:
            pass
        
        return None
    
    def get_form_metadata(self, form_id):
        """Get detailed form metadata"""
        try:
            query = f"{self.api_url}/systemforms({form_id})"
            params = {
                '$select': 'formid,name,description,type,objecttypecode,formactivationstate,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_optionset_metadata(self, optionset_id):
        """Get detailed option set metadata"""
        try:
            query = f"{self.api_url}/GlobalOptionSetDefinitions({optionset_id})"
            params = {
                '$select': 'MetadataId,Name,DisplayName,Description,IsGlobal'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                optionset = response.json()
                
                display_name = optionset.get('DisplayName', {})
                if isinstance(display_name, dict):
                    display_name = display_name.get('UserLocalizedLabel', {}).get('Label', optionset.get('Name', ''))
                
                description = optionset.get('Description', {})
                if isinstance(description, dict):
                    description = description.get('UserLocalizedLabel', {}).get('Label', '')
                
                return {
                    'MetadataId': optionset.get('MetadataId'),
                    'Name': optionset.get('Name'),
                    'DisplayName': display_name,
                    'Description': description,
                    'IsGlobal': optionset.get('IsGlobal', False)
                }
        except:
            pass
        
        return None
    
    def get_flow_metadata(self, flow_id):
        """Get detailed flow metadata"""
        if flow_id in self.all_flows:
            return self.all_flows[flow_id]
        
        try:
            query = f"{self.api_url}/workflows({flow_id})"
            params = {
                '$select': 'workflowid,name,uniquename,description,category,primaryentity,type,statecode,createdon,modifiedon,createdby,modifiedby'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                flow = response.json()
                self.all_flows[flow_id] = flow
                return flow
        except:
            pass
        
        return None
    
    def get_view_metadata(self, view_id):
        """Get detailed view metadata"""
        if view_id in self.all_views:
            return self.all_views[view_id]
        
        try:
            query = f"{self.api_url}/savedqueries({view_id})"
            params = {
                '$select': 'savedqueryid,name,description,returnedtypecode,querytype,isdefault,isquickfindquery,isprivate,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                view = response.json()
                self.all_views[view_id] = view
                return view
        except:
            pass
        
        return None
    
    def get_role_metadata(self, role_id):
        """Get detailed role metadata"""
        if role_id in self.all_roles:
            return self.all_roles[role_id]
        
        try:
            query = f"{self.api_url}/roles({role_id})"
            params = {
                '$select': 'roleid,name,businessunitid,iscustomizable,ismanaged,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                role = response.json()
                self.all_roles[role_id] = role
                return role
        except:
            pass
        
        return None
    
    def get_web_resource_metadata(self, wr_id):
        """Get detailed web resource metadata"""
        if wr_id in self.all_web_resources:
            return self.all_web_resources[wr_id]
        
        try:
            query = f"{self.api_url}/webresources({wr_id})"
            params = {
                '$select': 'webresourceid,name,displayname,description,webresourcetype,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                wr = response.json()
                self.all_web_resources[wr_id] = wr
                return wr
        except:
            pass
        
        return None
    
    def get_form_metadata(self, form_id):
        """Get detailed form metadata"""
        try:
            query = f"{self.api_url}/systemforms({form_id})"
            params = {
                '$select': 'formid,name,description,type,objecttypecode,formactivationstate,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_optionset_metadata(self, optionset_id):
        """Get detailed option set metadata"""
        try:
            query = f"{self.api_url}/GlobalOptionSetDefinitions({optionset_id})"
            params = {
                '$select': 'MetadataId,Name,DisplayName,Description,IsGlobal'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                optionset = response.json()
                
                display_name = optionset.get('DisplayName', {})
                if isinstance(display_name, dict):
                    display_name = display_name.get('UserLocalizedLabel', {}).get('Label', optionset.get('Name', ''))
                
                description = optionset.get('Description', {})
                if isinstance(description, dict):
                    description = description.get('UserLocalizedLabel', {}).get('Label', '')
                
                return {
                    'MetadataId': optionset.get('MetadataId'),
                    'Name': optionset.get('Name'),
                    'DisplayName': display_name,
                    'Description': description,
                    'IsGlobal': optionset.get('IsGlobal', False)
                }
        except:
            pass
        
        return None
    
    def get_attribute_metadata(self, entity_logical_name, attribute_id):
        """Get attribute/column metadata"""
        try:
            query = f"{self.api_url}/EntityDefinitions(LogicalName='{entity_logical_name}')/Attributes({attribute_id})"
            params = {
                '$select': 'LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                attr = response.json()
                
                display_name = attr.get('DisplayName', {})
                if isinstance(display_name, dict):
                    display_name = display_name.get('UserLocalizedLabel', {}).get('Label', attr.get('LogicalName', ''))
                
                return {
                    'LogicalName': attr.get('LogicalName'),
                    'SchemaName': attr.get('SchemaName'),
                    'DisplayName': display_name,
                    'AttributeType': attr.get('AttributeType'),
                    'IsCustomAttribute': attr.get('IsCustomAttribute', False)
                }
        except:
            pass
        
        return None
    
    def get_relationship_metadata(self, relationship_id):
        """Get relationship metadata"""
        try:
            # Try OneToMany first
            query = f"{self.api_url}/RelationshipDefinitions({relationship_id})"
            params = {
                '$select': 'SchemaName,RelationshipType'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                rel = response.json()
                return {
                    'SchemaName': rel.get('SchemaName'),
                    'RelationshipType': rel.get('RelationshipType')
                }
        except:
            pass
        
        return None
    
    def get_plugin_assembly_metadata(self, assembly_id):
        """Get plugin assembly metadata"""
        try:
            query = f"{self.api_url}/pluginassemblies({assembly_id})"
            params = {
                '$select': 'pluginassemblyid,name,version,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_plugin_type_metadata(self, plugin_type_id):
        """Get plugin type metadata"""
        try:
            query = f"{self.api_url}/plugintypes({plugin_type_id})"
            params = {
                '$select': 'plugintypeid,typename,friendlyname,name,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_sdk_message_step_metadata(self, step_id):
        """Get SDK message processing step metadata"""
        try:
            query = f"{self.api_url}/sdkmessageprocessingsteps({step_id})"
            params = {
                '$select': 'sdkmessageprocessingstepid,name,stage,mode,rank,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_canvas_app_metadata(self, app_id):
        """Get canvas app metadata"""
        try:
            query = f"{self.api_url}/canvasapps({app_id})"
            params = {
                '$select': 'canvasappid,name,displayname,description,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def get_environment_variable_definition_metadata(self, def_id):
        """Get environment variable definition metadata"""
        try:
            query = f"{self.api_url}/environmentvariabledefinitions({def_id})"
            params = {
                '$select': 'environmentvariabledefinitionid,schemaname,displayname,description,type,createdon,modifiedon'
            }
            
            response = requests.get(query, headers=self.get_headers(), params=params, timeout=10)
            
            if response.status_code == 200:
                return response.json()
        except:
            pass
        
        return None
    
    def create_individual_solution_workbook(self, solution, output_dir):
        """Create detailed workbook for a single solution"""
        components = self.all_components.get(solution['solution_id'], [])
        
        if len(components) == 0:
            return None
        
        # Create solution-specific folder
        safe_folder_name = "".join(c for c in solution['display_name'] if c.isalnum() or c in (' ', '-', '_')).strip()
        solution_folder = output_dir / safe_folder_name
        solution_folder.mkdir(exist_ok=True)
        
        # Create filename
        install_date = solution['installed_on'][:10] if solution['installed_on'] else 'Unknown'
        safe_name = "".join(c for c in solution['display_name'] if c.isalnum() or c in (' ', '-', '_')).strip()
        filename = f"{safe_name}_{install_date}.xlsx"
        filepath = solution_folder / filename
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        # Styles
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        # SHEET 1: Summary
        ws_summary = wb.create_sheet("Summary")
        ws_summary.append(['Property', 'Value'])
        
        for row in ws_summary['A1:B1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        ws_summary.append(['Solution Name', solution['display_name']])
        ws_summary.append(['Unique Name', solution['unique_name']])
        ws_summary.append(['Version', solution['version']])
        ws_summary.append(['Publisher', solution['publisher']])
        ws_summary.append(['Type', 'Managed' if solution['is_managed'] else 'Unmanaged'])
        ws_summary.append(['Installed On', install_date])
        ws_summary.append(['Description', solution['description']])
        ws_summary.append(['Total Components', len(components)])
        
        ws_summary.column_dimensions['A'].width = 30
        ws_summary.column_dimensions['B'].width = 50
        
        # SHEET 2: All Components Summary with Logical Names
        ws_all_components = wb.create_sheet("All Components Summary")
        ws_all_components.append(['Component Type', 'Type Code', 'Logical Name', 'Schema Name', 'Object ID', 'Component ID'])
        
        for row in ws_all_components['A1:F1']:
            for cell in row:
                cell.fill = header_fill
                cell.font = header_font
        
        print(f"      Fetching component names...", end=' ', flush=True)
        
        for comp in sorted(components, key=lambda c: (c['componenttype'], str(c['objectid']))):
            comp_type = self.COMPONENT_TYPES.get(comp['componenttype'], f"Type {comp['componenttype']}")
            
            # Get logical name and schema name for this component
            logical_name, schema_name = self.get_component_logical_name(comp['componenttype'], comp['objectid'])
            
            ws_all_components.append([
                comp_type,
                comp['componenttype'],
                logical_name,
                schema_name,
                str(comp['objectid']),
                str(comp['solutioncomponentid'])
            ])
        
        print("✓")
        
        for column in ws_all_components.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws_all_components.column_dimensions[column_letter].width = min(max_length + 2, 60)
        
        # Group components by type
        components_by_type = defaultdict(list)
        for comp in components:
            components_by_type[comp['componenttype']].append(comp)
        
        # Create a sheet for each component type
        for comp_type, comp_list in sorted(components_by_type.items()):
            comp_type_name = self.COMPONENT_TYPES.get(comp_type, f"Type {comp_type}")
            
            # Sanitize sheet name (max 31 chars, no special chars)
            sheet_name = f"{comp_type}-{comp_type_name}"[:31]
            sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_'))
            
            ws = wb.create_sheet(sheet_name)
            
            # Different columns based on component type
            if comp_type == 1:  # Tables/Entities
                ws.append(['Type Code', 'Table Name', 'Logical Name', 'Schema Name', 'Record Count', 'Is Custom', 'Primary ID', 'Primary Name', 'Object Type Code', 'Object ID'])
                
                for row in ws['A1:J1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    table_info = None
                    for tbl in self.all_tables.values():
                        if tbl.get('ObjectId') == comp['objectid']:
                            table_info = tbl
                            break
                    
                    if table_info:
                        record_count_display = f"{table_info['RecordCount']:,}" if table_info['RecordCount'] is not None else "N/A"
                        
                        ws.append([
                            comp_type,
                            table_info['DisplayName'],
                            table_info['LogicalName'],
                            table_info['SchemaName'],
                            record_count_display,
                            'Yes' if table_info['IsCustomEntity'] else 'No',
                            table_info['PrimaryIdAttribute'],
                            table_info['PrimaryNameAttribute'],
                            table_info['ObjectTypeCode'],
                            str(comp['objectid'])
                        ])
            
            elif comp_type == 29:  # Workflows/Flows
                ws.append(['Type Code', 'Flow Name', 'Unique Name', 'Description', 'Category', 'Primary Entity', 'Type', 'State', 'Created On', 'Modified On', 'Object ID'])
                
                for row in ws['A1:K1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    flow = self.get_flow_metadata(comp['objectid'])
                    
                    if flow:
                        ws.append([
                            comp_type,
                            flow.get('name', ''),
                            flow.get('uniquename', ''),
                            flow.get('description', ''),
                            flow.get('category', ''),
                            flow.get('primaryentity', ''),
                            flow.get('type', ''),
                            flow.get('statecode', ''),
                            flow.get('createdon', '')[:10] if flow.get('createdon') else '',
                            flow.get('modifiedon', '')[:10] if flow.get('modifiedon') else '',
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', '', '', '', '', '', '', '', str(comp['objectid'])])
            
            elif comp_type in [26, 10293]:  # Views/Saved Queries
                ws.append(['Type Code', 'View Name', 'Description', 'Entity', 'Query Type', 'Is Default', 'Is Quick Find', 'Is Private', 'Created On', 'Modified On', 'Object ID'])
                
                for row in ws['A1:K1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    view = self.get_view_metadata(comp['objectid'])
                    
                    if view:
                        ws.append([
                            comp_type,
                            view.get('name', ''),
                            view.get('description', ''),
                            view.get('returnedtypecode', ''),
                            view.get('querytype', ''),
                            view.get('isdefault', False),
                            view.get('isquickfindquery', False),
                            view.get('isprivate', False),
                            view.get('createdon', '')[:10] if view.get('createdon') else '',
                            view.get('modifiedon', '')[:10] if view.get('modifiedon') else '',
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', '', '', False, False, False, '', '', str(comp['objectid'])])
            
            elif comp_type == 20:  # Security Roles
                ws.append(['Type Code', 'Role Name', 'Business Unit ID', 'Is Customizable', 'Is Managed', 'Created On', 'Modified On', 'Object ID'])
                
                for row in ws['A1:H1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    role = self.get_role_metadata(comp['objectid'])
                    
                    if role:
                        ws.append([
                            comp_type,
                            role.get('name', ''),
                            str(role.get('_businessunitid_value', '')),
                            role.get('iscustomizable', {}).get('Value', False) if isinstance(role.get('iscustomizable'), dict) else role.get('iscustomizable', False),
                            role.get('ismanaged', False),
                            role.get('createdon', '')[:10] if role.get('createdon') else '',
                            role.get('modifiedon', '')[:10] if role.get('modifiedon') else '',
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', False, False, '', '', str(comp['objectid'])])
            
            elif comp_type == 61:  # Web Resources
                ws.append(['Type Code', 'Name', 'Display Name', 'Description', 'Type', 'Created On', 'Modified On', 'Object ID'])
                
                for row in ws['A1:H1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                web_resource_types = {
                    1: 'HTML', 2: 'CSS', 3: 'JavaScript', 4: 'XML', 5: 'PNG',
                    6: 'JPG', 7: 'GIF', 8: 'XAP', 9: 'XSL', 10: 'ICO', 11: 'SVG', 12: 'RESX'
                }
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    wr = self.get_web_resource_metadata(comp['objectid'])
                    
                    if wr:
                        wr_type = web_resource_types.get(wr.get('webresourcetype', 0), str(wr.get('webresourcetype', '')))
                        
                        ws.append([
                            comp_type,
                            wr.get('name', ''),
                            wr.get('displayname', ''),
                            wr.get('description', ''),
                            wr_type,
                            wr.get('createdon', '')[:10] if wr.get('createdon') else '',
                            wr.get('modifiedon', '')[:10] if wr.get('modifiedon') else '',
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', '', '', '', '', str(comp['objectid'])])
            
            elif comp_type == 60:  # System Forms
                ws.append(['Type Code', 'Form Name', 'Description', 'Type', 'Entity', 'State', 'Created On', 'Modified On', 'Object ID'])
                
                for row in ws['A1:I1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    form = self.get_form_metadata(comp['objectid'])
                    
                    if form:
                        ws.append([
                            comp_type,
                            form.get('name', ''),
                            form.get('description', ''),
                            form.get('type', ''),
                            form.get('objecttypecode', ''),
                            form.get('formactivationstate', ''),
                            form.get('createdon', '')[:10] if form.get('createdon') else '',
                            form.get('modifiedon', '')[:10] if form.get('modifiedon') else '',
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', '', '', '', '', '', str(comp['objectid'])])
            
            elif comp_type == 9:  # Option Sets
                ws.append(['Type Code', 'Name', 'Display Name', 'Description', 'Is Global', 'Object ID'])
                
                for row in ws['A1:F1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    optionset = self.get_optionset_metadata(comp['objectid'])
                    
                    if optionset:
                        ws.append([
                            comp_type,
                            optionset.get('Name', ''),
                            optionset.get('DisplayName', ''),
                            optionset.get('Description', ''),
                            optionset.get('IsGlobal', False),
                            str(comp['objectid'])
                        ])
                    else:
                        ws.append([comp_type, '', '', '', False, str(comp['objectid'])])
            
            else:  # All other component types - generic format
                ws.append(['Type Code', 'Component Type', 'Object ID', 'Component ID'])
                
                for row in ws['A1:D1']:
                    for cell in row:
                        cell.fill = header_fill
                        cell.font = header_font
                
                for comp in sorted(comp_list, key=lambda c: str(c['objectid'])):
                    ws.append([
                        comp_type,
                        comp_type_name,
                        str(comp['objectid']),
                        str(comp['solutioncomponentid'])
                    ])
            
            # Auto-size columns
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column_letter].width = min(max_length + 2, 60)
        
        # Save
        wb.save(filepath)
        return filepath
    
    def create_all_individual_workbooks(self, output_dir):
        """Create individual workbooks for each solution"""
        print("="*70)
        print("STEP 5: CREATING INDIVIDUAL SOLUTION WORKBOOKS")
        print("="*70 + "\n")
        
        created_files = []
        
        for i, solution in enumerate(self.all_solutions, 1):
            print(f"[{i}/{len(self.all_solutions)}] {solution['display_name']}...", end=' ', flush=True)
            
            filepath = self.create_individual_solution_workbook(solution, output_dir)
            
            if filepath:
                created_files.append(filepath)
                # Show folder/filename structure
                print(f"✓ {filepath.parent.name}/{filepath.name}")
            else:
                print(f"✗ No components")
        
        print(f"\n✓ Created {len(created_files)} individual solution workbooks\n")
        return created_files
    
    def run_full_analysis(self, output_dir, include_microsoft=False):
        """Run complete analysis"""
        
        # Step 1: Fetch solutions
        if self.fetch_all_solutions(include_microsoft) == 0:
            return None, []
        
        # Step 2: Collect all components
        self.collect_all_components()
        
        # Step 3: Analyze tables
        self.analyze_all_tables()
        
        # Step 4: Create master workbook
        master_file = self.create_master_workbook(output_dir)
        
        # Step 5: Create individual solution workbooks
        individual_files = self.create_all_individual_workbooks(output_dir)
        
        return master_file, individual_files


def main():
    parser = argparse.ArgumentParser(
        description='Export Power Platform solutions with master analysis workbook'
    )
    parser.add_argument('--url', required=True, help='Environment URL')
    parser.add_argument('--output-dir', default='solution_master_export', help='Output directory')
    parser.add_argument('--include-microsoft', action='store_true', help='Include Microsoft base solutions')
    parser.add_argument('--client-id', help='Azure AD Client ID (optional)')
    parser.add_argument('--tenant-id', help='Azure AD Tenant ID (optional)')
    
    args = parser.parse_args()
    
    print("\n" + "="*70)
    print("POWER PLATFORM MASTER SOLUTION ANALYZER")
    print("Collecting ALL data in memory for cross-solution analysis")
    print("="*70)
    
    analyzer = MasterSolutionAnalyzer(args.url)
    
    # Authenticate
    if not analyzer.authenticate_interactive(args.client_id, args.tenant_id):
        print("Authentication failed. Exiting.")
        sys.exit(1)
    
    # Create output directory with environment name and date
    extract_date = datetime.now().strftime('%Y%m%d')
    folder_name = f"Environment Solution Analysis {analyzer.environment_name} {extract_date}"
    output_path = Path(args.output_dir) / folder_name
    output_path.mkdir(parents=True, exist_ok=True)
    
    print(f"\nOutput folder: {output_path}\n")
    
    # Run full analysis
    master_file, individual_files = analyzer.run_full_analysis(output_path, args.include_microsoft)
    
    if master_file:
        print("="*70)
        print(f"✓ ANALYSIS COMPLETE")
        print(f"  Environment: {analyzer.environment_name}")
        print(f"  Extract Date: {extract_date}")
        print(f"  Master workbook: {master_file.name}")
        print(f"  Individual workbooks: {len(individual_files)}")
        print(f"  Output folder: {output_path}")
        print("="*70 + "\n")
    else:
        print("No solutions found to analyze.")
        sys.exit(1)


if __name__ == "__main__":
    main()
